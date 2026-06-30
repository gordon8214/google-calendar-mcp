import { CallToolResult, McpError, ErrorCode } from "@modelcontextprotocol/sdk/types.js";
import { OAuth2Client } from "google-auth-library";
import { calendar_v3 } from "googleapis";
import { randomUUID } from "crypto";
import { BaseToolHandler } from "./BaseToolHandler.js";
import { MoveEventInput } from "../../tools/registry.js";
import { MoveEventResponse } from "../../types/structured-responses.js";
import { convertGoogleEventToStructured } from "../../types/structured-responses.js";
import { createStructuredResponse } from "../../utils/response-builder.js";

/**
 * Moves an event from one calendar to another.
 *
 * Two strategies, selected automatically:
 *  - Same account: Google's native `events.move`. One call; the event ID, attendees,
 *    organizer, and RSVPs are all preserved.
 *  - Different accounts: Google's API cannot move across accounts, so we copy the event
 *    into the destination (`events.insert`) and then delete the original (`events.delete`).
 *    The copy gets a new event ID and the destination account becomes the organizer.
 *
 * The copy-then-delete ordering is deliberate: if the copy fails the source is untouched,
 * and if the delete fails after a successful copy we keep the copy and warn rather than
 * throwing — the user's data is never lost, at worst it is briefly duplicated.
 */
export class MoveEventHandler extends BaseToolHandler {
    async runTool(args: any, accounts: Map<string, OAuth2Client>): Promise<CallToolResult> {
        const validArgs = args as MoveEventInput;

        // Resolve the source: account (write access) + calendar name -> ID + Calendar API client.
        const {
            calendar: sourceCalendar,
            accountId: sourceAccountId,
            calendarId: sourceCalendarId
        } = await this.setupOperation(validArgs.account, validArgs.calendarId, accounts, 'write');

        // Resolve the destination. When no destination account is given, pin it to the RESOLVED
        // source account (not the raw arg) so that omitting destinationAccount reliably means "a
        // move between two calendars on the same account". Defaulting to the raw validArgs.account
        // would, when both account args are omitted, let the destination auto-select a different
        // account and silently downgrade a lossless native move into a lossy copy-delete.
        const {
            client: destClient,
            accountId: destAccountId,
            calendarId: destCalendarId
        } = await this.getClientWithAutoSelection(
            validArgs.destinationAccount ?? sourceAccountId,
            validArgs.destinationCalendarId,
            accounts,
            'write'
        );

        // No-op guard, now that names are resolved to concrete IDs (the schema-level refine only
        // catches literal string matches before resolution).
        if (sourceAccountId === destAccountId && sourceCalendarId === destCalendarId) {
            throw new McpError(
                ErrorCode.InvalidRequest,
                `Source and destination resolve to the same calendar ("${destCalendarId}" on account "${destAccountId}"); nothing to move.`
            );
        }

        // Fetch the full source event (no field mask) — needed for the recurring-instance guard
        // and, on the cross-account path, to build the copy.
        let sourceEvent: calendar_v3.Schema$Event;
        try {
            const response = await sourceCalendar.events.get({
                calendarId: sourceCalendarId,
                eventId: validArgs.eventId
            });
            sourceEvent = response.data;
        } catch (error) {
            throw this.handleGoogleApiError(error);
        }

        // A single instance of a recurring series cannot be moved: the native API rejects it, and
        // a cross-account copy would silently detach the instance from its series. Guard both paths.
        if (sourceEvent.recurringEventId) {
            throw new McpError(
                ErrorCode.InvalidRequest,
                "Cannot move a single instance of a recurring event. Move the entire series instead, " +
                "or use update-event to reschedule this occurrence."
            );
        }

        const sendUpdates = validArgs.sendUpdates;

        // ---- Same account: native move ----
        if (sourceAccountId === destAccountId) {
            let movedEvent: calendar_v3.Schema$Event;
            try {
                const response = await sourceCalendar.events.move({
                    calendarId: sourceCalendarId,
                    eventId: validArgs.eventId,
                    destination: destCalendarId,
                    sendUpdates
                });
                movedEvent = response.data;
            } catch (error) {
                throw this.handleGoogleApiError(error);
            }

            const response: MoveEventResponse = {
                event: convertGoogleEventToStructured(movedEvent, destCalendarId, destAccountId),
                moved: true,
                method: 'native',
                source: {
                    accountId: sourceAccountId,
                    calendarId: sourceCalendarId,
                    eventId: validArgs.eventId,
                    deleted: true
                },
                destination: {
                    accountId: destAccountId,
                    calendarId: destCalendarId,
                    eventId: movedEvent.id ?? validArgs.eventId
                }
            };
            return createStructuredResponse(response);
        }

        // ---- Different accounts: copy into destination, then delete from source ----
        const { body, warnings } = this.buildCrossAccountCopy(sourceEvent, {
            copyAttendees: validArgs.copyAttendees ?? false,
            recreateConference: validArgs.recreateConference ?? false
        });

        let insertedEvent: calendar_v3.Schema$Event;
        try {
            const destCalendar = this.getCalendar(destClient);
            const response = await destCalendar.events.insert({
                calendarId: destCalendarId,
                requestBody: body,
                sendUpdates,
                ...(body.attachments ? { supportsAttachments: true } : {}),
                ...(body.conferenceData ? { conferenceDataVersion: 1 } : {})
            });
            insertedEvent = response.data;
        } catch (error) {
            // Insert failed — the source event is untouched, so surface the error directly.
            throw this.handleGoogleApiError(error);
        }

        // Copy succeeded. Remove the original; if this fails, keep the copy and warn (the data is
        // safe, just duplicated) instead of throwing.
        let sourceDeleted = true;
        try {
            await sourceCalendar.events.delete({
                calendarId: sourceCalendarId,
                eventId: validArgs.eventId,
                sendUpdates
            });
        } catch (error) {
            sourceDeleted = false;
            warnings.push(
                `The event was copied to "${destCalendarId}" on account "${destAccountId}" as "${insertedEvent.id}", ` +
                `but the original could not be removed from "${sourceCalendarId}" on account "${sourceAccountId}" ` +
                `(${this.formatGoogleApiError(error)}). Delete the original manually to avoid a duplicate.`
            );
        }

        const response: MoveEventResponse = {
            event: convertGoogleEventToStructured(insertedEvent, destCalendarId, destAccountId),
            moved: true,
            method: 'copy-delete',
            source: {
                accountId: sourceAccountId,
                calendarId: sourceCalendarId,
                eventId: validArgs.eventId,
                deleted: sourceDeleted
            },
            destination: {
                accountId: destAccountId,
                calendarId: destCalendarId,
                eventId: insertedEvent.id ?? ''
            },
            ...(warnings.length > 0 ? { warnings } : {})
        };
        return createStructuredResponse(response);
    }

    /**
     * Builds the request body for a cross-account copy from a fetched source event.
     *
     * Account-bound and read-only fields (id, etag, iCalUID, organizer, creator, html/hangout
     * links, conferenceData, recurringEventId, sequence, status, etc.) are intentionally NOT
     * carried over — they either error on insert or are meaningless in the destination account.
     * Only portable user content and settings are copied. Returns the body plus any warnings
     * describing what could not be transferred faithfully.
     */
    private buildCrossAccountCopy(
        source: calendar_v3.Schema$Event,
        options: { copyAttendees: boolean; recreateConference: boolean }
    ): { body: calendar_v3.Schema$Event; warnings: string[] } {
        const warnings: string[] = [];
        const body: calendar_v3.Schema$Event = {};

        // User content
        if (source.summary != null) body.summary = source.summary;
        if (source.description != null) body.description = source.description;
        if (source.location != null) body.location = source.location;
        if (source.start) body.start = source.start;
        if (source.end) body.end = source.end;

        // Portable settings
        if (source.colorId != null) body.colorId = source.colorId;
        if (source.reminders) body.reminders = source.reminders;
        if (source.transparency != null) body.transparency = source.transparency;
        if (source.visibility != null) body.visibility = source.visibility;
        if (source.extendedProperties) body.extendedProperties = source.extendedProperties;
        if (source.source) body.source = source.source;
        if (source.guestsCanInviteOthers != null) body.guestsCanInviteOthers = source.guestsCanInviteOthers;
        if (source.guestsCanModify != null) body.guestsCanModify = source.guestsCanModify;
        if (source.guestsCanSeeOtherGuests != null) body.guestsCanSeeOtherGuests = source.guestsCanSeeOtherGuests;
        if (source.anyoneCanAddSelf != null) body.anyoneCanAddSelf = source.anyoneCanAddSelf;

        // Only the default event type is portable; Workspace-only types (outOfOffice, focusTime,
        // workingLocation) are tied to a primary calendar and would error elsewhere.
        if (source.eventType && source.eventType !== 'default') {
            warnings.push(
                `Event type "${source.eventType}" is not portable across accounts; the moved event was created as a default event.`
            );
        }

        // Attachments are copied by reference; the fileUrl/fileId point at the source account's
        // Drive, so they may not be openable from the destination account.
        if (source.attachments && source.attachments.length > 0) {
            body.attachments = source.attachments;
            warnings.push(
                "Attachments were copied by reference and point at the source account's Drive; they may not be accessible from the destination account."
            );
        }

        // Recurring series master: copy the rules, but exceptions/history cannot follow.
        if (source.recurrence && source.recurrence.length > 0) {
            body.recurrence = source.recurrence;
            warnings.push(
                "Recurrence rules were copied, but per-instance changes (moved, edited, or deleted occurrences) and event history were not transferred."
            );
        }

        // Attendees: stripped by default (the move becomes a clean personal copy). With
        // copyAttendees, re-invite from the destination account, which resets RSVPs.
        if (source.attendees && source.attendees.length > 0) {
            if (options.copyAttendees) {
                body.attendees = source.attendees
                    .filter(a => a.email)
                    .map(a => ({
                        email: a.email!,
                        displayName: a.displayName ?? undefined,
                        optional: a.optional ?? undefined,
                        comment: a.comment ?? undefined,
                        resource: a.resource ?? undefined
                    }));
                warnings.push(
                    "Attendees were re-invited from the destination account; their previous RSVPs were reset."
                );
            } else {
                warnings.push(
                    "Attendees were not copied to the destination event. Set copyAttendees to re-invite them."
                );
            }
        }

        // Google Meet / conference data is bound to the original organizer's account and cannot
        // be transferred. Optionally mint a fresh link in the destination.
        if (source.conferenceData || source.hangoutLink) {
            if (options.recreateConference) {
                body.conferenceData = {
                    createRequest: {
                        requestId: randomUUID(),
                        conferenceSolutionKey: { type: 'hangoutsMeet' }
                    }
                };
                warnings.push(
                    "A new Google Meet link was created in the destination; it differs from the original link."
                );
            } else {
                warnings.push(
                    "The original Google Meet link does not transfer between accounts and was removed."
                );
            }
        }

        return { body, warnings };
    }
}
