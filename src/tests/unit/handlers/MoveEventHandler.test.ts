import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MoveEventHandler } from '../../../handlers/core/MoveEventHandler.js';
import { OAuth2Client } from 'google-auth-library';
import { CalendarRegistry } from '../../../services/CalendarRegistry.js';

// Mock the googleapis module (the handler's account/calendar resolution is spied below, so the
// real google.calendar is never invoked — this just keeps the import resolvable).
vi.mock('googleapis', () => ({
  google: {
    calendar: vi.fn(() => ({
      events: { get: vi.fn(), move: vi.fn(), insert: vi.fn(), delete: vi.fn() }
    }))
  },
  calendar_v3: {}
}));

interface MockCtx {
  handler: MoveEventHandler;
  sourceCalendar: any;
  destCalendar: any;
  mockAccounts: Map<string, OAuth2Client>;
}

/**
 * Wires up a MoveEventHandler with mocked source/destination resolution.
 * setupOperation resolves the SOURCE; getClientWithAutoSelection resolves the DESTINATION;
 * getCalendar returns the destination Calendar API (used only on the cross-account path).
 */
function setup(opts: {
  sourceAccountId: string;
  sourceCalendarId: string;
  destAccountId: string;
  destCalendarId: string;
}): MockCtx {
  const handler = new MoveEventHandler();
  const srcClient = new OAuth2Client();
  const destClient = new OAuth2Client();

  const sourceCalendar = { events: { get: vi.fn(), move: vi.fn(), delete: vi.fn() } };
  const destCalendar = { events: { insert: vi.fn() } };

  vi.spyOn(handler as any, 'setupOperation').mockResolvedValue({
    client: srcClient,
    calendar: sourceCalendar,
    accountId: opts.sourceAccountId,
    calendarId: opts.sourceCalendarId
  });
  vi.spyOn(handler as any, 'getClientWithAutoSelection').mockResolvedValue({
    client: destClient,
    accountId: opts.destAccountId,
    calendarId: opts.destCalendarId,
    wasAutoSelected: false
  });
  vi.spyOn(handler as any, 'getCalendar').mockReturnValue(destCalendar);

  return {
    handler,
    sourceCalendar,
    destCalendar,
    mockAccounts: new Map([
      [opts.sourceAccountId, srcClient],
      [opts.destAccountId, destClient]
    ])
  };
}

beforeEach(() => {
  CalendarRegistry.resetInstance();
});

describe('MoveEventHandler - same account (native move)', () => {
  it('uses events.move and preserves the event ID', async () => {
    const ctx = setup({
      sourceAccountId: 'gordon',
      sourceCalendarId: 'primary',
      destAccountId: 'gordon',
      destCalendarId: 'work@group.calendar.google.com'
    });
    ctx.sourceCalendar.events.get.mockResolvedValue({ data: { id: 'event123', summary: 'Standup' } });
    ctx.sourceCalendar.events.move.mockResolvedValue({ data: { id: 'event123', summary: 'Standup' } });

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'work@group.calendar.google.com',
      sendUpdates: 'none'
    }, ctx.mockAccounts);

    expect(ctx.sourceCalendar.events.move).toHaveBeenCalledWith({
      calendarId: 'primary',
      eventId: 'event123',
      destination: 'work@group.calendar.google.com',
      sendUpdates: 'none'
    });
    expect(ctx.destCalendar.events.insert).not.toHaveBeenCalled();
    expect(ctx.sourceCalendar.events.delete).not.toHaveBeenCalled();

    const response = JSON.parse(result.content[0].text as string);
    expect(response.method).toBe('native');
    expect(response.moved).toBe(true);
    expect(response.source).toMatchObject({ eventId: 'event123', deleted: true });
    expect(response.destination.eventId).toBe('event123');
    expect(response.warnings).toBeUndefined();
  });

  it('rejects when source and destination resolve to the same calendar', async () => {
    const ctx = setup({
      sourceAccountId: 'gordon',
      sourceCalendarId: 'primary',
      destAccountId: 'gordon',
      destCalendarId: 'primary'
    });

    await expect(ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary'
    }, ctx.mockAccounts)).rejects.toThrow(/same calendar/i);

    expect(ctx.sourceCalendar.events.get).not.toHaveBeenCalled();
    expect(ctx.sourceCalendar.events.move).not.toHaveBeenCalled();
  });

  it('allows moving a recurring series master (recurrence array, no recurringEventId)', async () => {
    const ctx = setup({
      sourceAccountId: 'gordon',
      sourceCalendarId: 'primary',
      destAccountId: 'gordon',
      destCalendarId: 'work@group.calendar.google.com'
    });
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: { id: 'series123', recurrence: ['RRULE:FREQ=WEEKLY;COUNT=10'] }
    });
    ctx.sourceCalendar.events.move.mockResolvedValue({ data: { id: 'series123' } });

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'series123',
      destinationCalendarId: 'work@group.calendar.google.com'
    }, ctx.mockAccounts);

    expect(ctx.sourceCalendar.events.move).toHaveBeenCalled();
    const response = JSON.parse(result.content[0].text as string);
    expect(response.method).toBe('native');
  });

  it('pins the destination to the resolved source account when destinationAccount is omitted', async () => {
    // Regression guard: the destination must default to the RESOLVED source account, not the raw
    // (here undefined) account arg, so omitting both account args stays a same-account native move
    // rather than silently auto-selecting a different account and copy-deleting.
    const ctx = setup({
      sourceAccountId: 'gordon',
      sourceCalendarId: 'primary',
      destAccountId: 'gordon',
      destCalendarId: 'work@group.calendar.google.com'
    });
    ctx.sourceCalendar.events.get.mockResolvedValue({ data: { id: 'event123' } });
    ctx.sourceCalendar.events.move.mockResolvedValue({ data: { id: 'event123' } });

    await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'work@group.calendar.google.com'
      // account and destinationAccount both omitted
    }, ctx.mockAccounts);

    expect((ctx.handler as any).getClientWithAutoSelection).toHaveBeenCalledWith(
      'gordon',
      'work@group.calendar.google.com',
      ctx.mockAccounts,
      'write'
    );
    expect(ctx.destCalendar.events.insert).not.toHaveBeenCalled();
  });
});

describe('MoveEventHandler - cross account (copy + delete)', () => {
  const crossOpts = {
    sourceAccountId: 'gordon',
    sourceCalendarId: 'primary',
    destAccountId: 'hability',
    destCalendarId: 'primary'
  };

  it('copies portable fields, strips account-bound fields, then deletes the original', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        etag: '"etag"',
        iCalUID: 'abc@google.com',
        summary: 'Project sync',
        description: 'notes',
        location: 'Room 1',
        start: { dateTime: '2026-07-01T10:00:00-07:00', timeZone: 'America/Los_Angeles' },
        end: { dateTime: '2026-07-01T11:00:00-07:00', timeZone: 'America/Los_Angeles' },
        organizer: { email: 'gordon@example.com', self: true },
        creator: { email: 'gordon@example.com' },
        htmlLink: 'https://calendar.google.com/event?eid=xyz',
        sequence: 3,
        attendees: [{ email: 'a@example.com', responseStatus: 'accepted' }]
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456', summary: 'Project sync' } });
    ctx.sourceCalendar.events.delete.mockResolvedValue({ data: {} });

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability',
      account: 'gordon',
      sendUpdates: 'none'
    }, ctx.mockAccounts);

    expect(ctx.sourceCalendar.events.move).not.toHaveBeenCalled();

    const insertArgs = ctx.destCalendar.events.insert.mock.calls[0][0];
    expect(insertArgs.calendarId).toBe('primary');
    expect(insertArgs.requestBody.summary).toBe('Project sync');
    expect(insertArgs.requestBody.start).toEqual({ dateTime: '2026-07-01T10:00:00-07:00', timeZone: 'America/Los_Angeles' });
    // Account-bound / read-only fields must NOT be carried over.
    expect(insertArgs.requestBody.id).toBeUndefined();
    expect(insertArgs.requestBody.etag).toBeUndefined();
    expect(insertArgs.requestBody.iCalUID).toBeUndefined();
    expect(insertArgs.requestBody.organizer).toBeUndefined();
    expect(insertArgs.requestBody.creator).toBeUndefined();
    expect(insertArgs.requestBody.htmlLink).toBeUndefined();
    expect(insertArgs.requestBody.sequence).toBeUndefined();
    // Attendees stripped by default.
    expect(insertArgs.requestBody.attendees).toBeUndefined();

    expect(ctx.sourceCalendar.events.delete).toHaveBeenCalledWith({
      calendarId: 'primary',
      eventId: 'event123',
      sendUpdates: 'none'
    });

    const response = JSON.parse(result.content[0].text as string);
    expect(response.method).toBe('copy-delete');
    expect(response.moved).toBe(true);
    expect(response.source).toMatchObject({ eventId: 'event123', deleted: true });
    expect(response.destination.eventId).toBe('newEvent456');
    expect(response.warnings).toContain('Attendees were not copied to the destination event. Set copyAttendees to re-invite them.');
  });

  it('carries all-day date fields verbatim', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        summary: 'Holiday',
        start: { date: '2026-07-01' },
        end: { date: '2026-07-02' }
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456' } });
    ctx.sourceCalendar.events.delete.mockResolvedValue({ data: {} });

    await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability'
    }, ctx.mockAccounts);

    const insertArgs = ctx.destCalendar.events.insert.mock.calls[0][0];
    expect(insertArgs.requestBody.start).toEqual({ date: '2026-07-01' });
    expect(insertArgs.requestBody.end).toEqual({ date: '2026-07-02' });
  });

  it('re-invites attendees when copyAttendees is true, stripping per-account sub-fields', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        summary: 'Team meeting',
        start: { dateTime: '2026-07-01T10:00:00Z' },
        end: { dateTime: '2026-07-01T11:00:00Z' },
        attendees: [
          { email: 'a@example.com', responseStatus: 'accepted', organizer: true, self: true },
          { email: 'b@example.com', responseStatus: 'declined', optional: true },
          { displayName: 'No-email entry', responseStatus: 'needsAction' } // must be filtered out
        ]
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456' } });
    ctx.sourceCalendar.events.delete.mockResolvedValue({ data: {} });

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability',
      copyAttendees: true
    }, ctx.mockAccounts);

    const insertArgs = ctx.destCalendar.events.insert.mock.calls[0][0];
    expect(insertArgs.requestBody.attendees).toEqual([
      { email: 'a@example.com', displayName: undefined, optional: undefined, comment: undefined, resource: undefined },
      { email: 'b@example.com', displayName: undefined, optional: true, comment: undefined, resource: undefined }
    ]);

    const response = JSON.parse(result.content[0].text as string);
    expect(response.warnings).toContain('Attendees were re-invited from the destination account; their previous RSVPs were reset.');
  });

  it('drops the Google Meet link by default and warns', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        summary: 'Call',
        start: { dateTime: '2026-07-01T10:00:00Z' },
        end: { dateTime: '2026-07-01T11:00:00Z' },
        hangoutLink: 'https://meet.google.com/abc-defg-hij',
        conferenceData: { conferenceId: 'abc-defg-hij' }
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456' } });
    ctx.sourceCalendar.events.delete.mockResolvedValue({ data: {} });

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability'
    }, ctx.mockAccounts);

    const insertArgs = ctx.destCalendar.events.insert.mock.calls[0][0];
    expect(insertArgs.requestBody.conferenceData).toBeUndefined();
    expect(insertArgs.conferenceDataVersion).toBeUndefined();

    const response = JSON.parse(result.content[0].text as string);
    expect(response.warnings).toContain('The original Google Meet link does not transfer between accounts and was removed.');
  });

  it('mints a new Meet link when recreateConference is true', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        summary: 'Call',
        start: { dateTime: '2026-07-01T10:00:00Z' },
        end: { dateTime: '2026-07-01T11:00:00Z' },
        hangoutLink: 'https://meet.google.com/abc-defg-hij'
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456' } });
    ctx.sourceCalendar.events.delete.mockResolvedValue({ data: {} });

    await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability',
      recreateConference: true
    }, ctx.mockAccounts);

    const insertArgs = ctx.destCalendar.events.insert.mock.calls[0][0];
    expect(insertArgs.conferenceDataVersion).toBe(1);
    expect(insertArgs.requestBody.conferenceData.createRequest.conferenceSolutionKey.type).toBe('hangoutsMeet');
    expect(insertArgs.requestBody.conferenceData.createRequest.requestId).toBeTruthy();
  });

  it('copies recurrence rules for a series master and warns about exceptions', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        summary: 'Weekly sync',
        start: { dateTime: '2026-07-01T10:00:00Z' },
        end: { dateTime: '2026-07-01T11:00:00Z' },
        recurrence: ['RRULE:FREQ=WEEKLY;COUNT=10']
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456' } });
    ctx.sourceCalendar.events.delete.mockResolvedValue({ data: {} });

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability'
    }, ctx.mockAccounts);

    const insertArgs = ctx.destCalendar.events.insert.mock.calls[0][0];
    expect(insertArgs.requestBody.recurrence).toEqual(['RRULE:FREQ=WEEKLY;COUNT=10']);

    const response = JSON.parse(result.content[0].text as string);
    expect(response.warnings.some((w: string) => /Recurrence rules were copied/.test(w))).toBe(true);
  });

  it('keeps the copy and warns (no throw) when deleting the original fails', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: {
        id: 'event123',
        summary: 'Standup',
        start: { dateTime: '2026-07-01T10:00:00Z' },
        end: { dateTime: '2026-07-01T11:00:00Z' }
      }
    });
    ctx.destCalendar.events.insert.mockResolvedValue({ data: { id: 'newEvent456' } });
    ctx.sourceCalendar.events.delete.mockRejectedValue(new Error('insufficient permissions'));

    const result = await ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability'
    }, ctx.mockAccounts);

    const response = JSON.parse(result.content[0].text as string);
    expect(response.moved).toBe(true);
    expect(response.method).toBe('copy-delete');
    expect(response.source.deleted).toBe(false);
    expect(response.destination.eventId).toBe('newEvent456');
    expect(response.warnings.some((w: string) => /could not be removed/.test(w))).toBe(true);
  });

  it('surfaces the insert error and leaves the source intact when the copy fails', async () => {
    const ctx = setup(crossOpts);
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: { id: 'event123', summary: 'Standup', start: {}, end: {} }
    });
    ctx.destCalendar.events.insert.mockRejectedValue(new Error('Bad Request'));

    await expect(ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123',
      destinationCalendarId: 'primary',
      destinationAccount: 'hability'
    }, ctx.mockAccounts)).rejects.toThrow();

    expect(ctx.sourceCalendar.events.delete).not.toHaveBeenCalled();
  });
});

describe('MoveEventHandler - recurring instance guard', () => {
  it('rejects moving a single instance of a recurring series (both paths)', async () => {
    const ctx = setup({
      sourceAccountId: 'gordon',
      sourceCalendarId: 'primary',
      destAccountId: 'gordon',
      destCalendarId: 'work@group.calendar.google.com'
    });
    ctx.sourceCalendar.events.get.mockResolvedValue({
      data: { id: 'event123_20260701T100000Z', recurringEventId: 'master123' }
    });

    await expect(ctx.handler.runTool({
      calendarId: 'primary',
      eventId: 'event123_20260701T100000Z',
      destinationCalendarId: 'work@group.calendar.google.com'
    }, ctx.mockAccounts)).rejects.toThrow(/single instance/i);

    expect(ctx.sourceCalendar.events.move).not.toHaveBeenCalled();
    expect(ctx.destCalendar.events.insert).not.toHaveBeenCalled();
  });
});
