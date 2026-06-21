import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { OAuth2Client } from "google-auth-library";
import { calendar_v3 } from "googleapis";
import { BaseToolHandler } from "./BaseToolHandler.js";
import { UpdateCalendarInput } from "../../tools/registry.js";
import { UpdateCalendarResponse, convertCalendarToStructured } from "../../types/structured-responses.js";
import { createStructuredResponse } from "../../utils/response-builder.js";

export class UpdateCalendarHandler extends BaseToolHandler {
    async runTool(args: any, accounts: Map<string, OAuth2Client>): Promise<CallToolResult> {
        const validArgs = args as UpdateCalendarInput;

        // Resolve account (write access) and calendar name -> ID in one call
        const { client, calendar, calendarId } = await this.setupOperation(
            validArgs.account,
            validArgs.calendarId,
            accounts,
            'write'
        );

        // Partition the requested changes across the two Google API resources:
        //   - `calendars.patch`     => shared, intrinsic calendar fields (requires owner access)
        //   - `calendarList.patch`  => the signed-in user's personal view/overrides
        const updated: string[] = [];

        const calendarBody: calendar_v3.Schema$Calendar = {};
        if (validArgs.summary !== undefined) { calendarBody.summary = validArgs.summary; updated.push('summary'); }
        if (validArgs.description !== undefined) { calendarBody.description = validArgs.description; updated.push('description'); }
        if (validArgs.location !== undefined) { calendarBody.location = validArgs.location; updated.push('location'); }
        if (validArgs.timeZone !== undefined) { calendarBody.timeZone = validArgs.timeZone; updated.push('timeZone'); }

        const calendarListBody: calendar_v3.Schema$CalendarListEntry = {};
        if (validArgs.summaryOverride !== undefined) { calendarListBody.summaryOverride = validArgs.summaryOverride; updated.push('summaryOverride'); }
        if (validArgs.colorId !== undefined) { calendarListBody.colorId = validArgs.colorId; updated.push('colorId'); }
        if (validArgs.backgroundColor !== undefined) { calendarListBody.backgroundColor = validArgs.backgroundColor; updated.push('backgroundColor'); }
        if (validArgs.foregroundColor !== undefined) { calendarListBody.foregroundColor = validArgs.foregroundColor; updated.push('foregroundColor'); }
        if (validArgs.hidden !== undefined) { calendarListBody.hidden = validArgs.hidden; updated.push('hidden'); }
        if (validArgs.selected !== undefined) { calendarListBody.selected = validArgs.selected; updated.push('selected'); }
        if (validArgs.defaultReminders !== undefined) { calendarListBody.defaultReminders = validArgs.defaultReminders; updated.push('defaultReminders'); }
        if (validArgs.notificationSettings !== undefined) { calendarListBody.notificationSettings = validArgs.notificationSettings; updated.push('notificationSettings'); }

        try {
            if (Object.keys(calendarBody).length > 0) {
                await calendar.calendars.patch({ calendarId, requestBody: calendarBody });
            }

            if (Object.keys(calendarListBody).length > 0) {
                // colorRgbFormat is required for custom hex background/foreground colors to take effect
                const usesCustomColors = validArgs.backgroundColor !== undefined || validArgs.foregroundColor !== undefined;
                await calendar.calendarList.patch({
                    calendarId,
                    requestBody: calendarListBody,
                    colorRgbFormat: usesCustomColors ? true : undefined
                });
            }
        } catch (error) {
            throw this.handleGoogleApiError(error);
        }

        // Calendar metadata changed; clear the registry cache (5-min TTL) so subsequent
        // list-calendars / name resolution reflects the update.
        this.calendarRegistry.clearCache();

        // Re-read the unified calendarList entry to return the fresh, merged state
        const entry = await this.getCalendarDetails(client, calendarId);

        const response: UpdateCalendarResponse = {
            calendar: convertCalendarToStructured(entry, calendarId),
            updated
        };

        return createStructuredResponse(response);
    }
}
