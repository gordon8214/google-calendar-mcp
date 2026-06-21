import { describe, it, expect, vi, beforeEach } from 'vitest';
import { UpdateCalendarHandler } from '../../../handlers/core/UpdateCalendarHandler.js';
import { OAuth2Client } from 'google-auth-library';
import { CalendarRegistry } from '../../../services/CalendarRegistry.js';

// Mock the googleapis module
vi.mock('googleapis', () => ({
  google: {
    calendar: vi.fn(() => ({
      calendars: { patch: vi.fn() },
      calendarList: { patch: vi.fn(), get: vi.fn() }
    }))
  },
  calendar_v3: {}
}));

describe('UpdateCalendarHandler', () => {
  let handler: UpdateCalendarHandler;
  let mockOAuth2Client: OAuth2Client;
  let mockAccounts: Map<string, OAuth2Client>;
  let mockCalendar: any;

  beforeEach(() => {
    CalendarRegistry.resetInstance();

    handler = new UpdateCalendarHandler();
    mockOAuth2Client = new OAuth2Client();
    mockAccounts = new Map([['test', mockOAuth2Client]]);

    mockCalendar = {
      calendars: { patch: vi.fn().mockResolvedValue({ data: {} }) },
      calendarList: {
        patch: vi.fn().mockResolvedValue({ data: {} }),
        // getCalendarDetails re-reads the entry after patching
        get: vi.fn().mockResolvedValue({
          data: {
            id: 'cal@example.com',
            summary: 'My Calendar',
            timeZone: 'America/New_York',
            colorId: '5',
            backgroundColor: '#0088aa',
            selected: true
          }
        })
      }
    };

    vi.spyOn(handler as any, 'getCalendar').mockReturnValue(mockCalendar);
    vi.spyOn(handler as any, 'getClientWithAutoSelection').mockResolvedValue({
      client: mockOAuth2Client,
      accountId: 'test',
      calendarId: 'cal@example.com',
      wasAutoSelected: true
    });
  });

  describe('Field routing', () => {
    it('routes shared calendar fields to calendars.patch only', async () => {
      const args = {
        calendarId: 'cal@example.com',
        timeZone: 'America/New_York',
        summary: 'New Name',
        description: 'Desc',
        location: 'Earth'
      };

      const result = await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendars.patch).toHaveBeenCalledWith({
        calendarId: 'cal@example.com',
        requestBody: {
          summary: 'New Name',
          description: 'Desc',
          location: 'Earth',
          timeZone: 'America/New_York'
        }
      });
      expect(mockCalendar.calendarList.patch).not.toHaveBeenCalled();

      const response = JSON.parse(result.content[0].text);
      expect(response.updated).toEqual(['summary', 'description', 'location', 'timeZone']);
      expect(response.calendar.id).toBe('cal@example.com');
    });

    it('routes per-user fields to calendarList.patch only', async () => {
      const args = {
        calendarId: 'cal@example.com',
        colorId: '5',
        summaryOverride: 'My nickname',
        hidden: false,
        selected: true
      };

      const result = await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendars.patch).not.toHaveBeenCalled();
      expect(mockCalendar.calendarList.patch).toHaveBeenCalledWith({
        calendarId: 'cal@example.com',
        requestBody: {
          summaryOverride: 'My nickname',
          colorId: '5',
          hidden: false,
          selected: true
        },
        colorRgbFormat: undefined
      });

      const response = JSON.parse(result.content[0].text);
      expect(response.updated).toEqual(['summaryOverride', 'colorId', 'hidden', 'selected']);
    });

    it('sets colorRgbFormat=true when custom hex colors are provided', async () => {
      const args = {
        calendarId: 'cal@example.com',
        backgroundColor: '#0088aa',
        foregroundColor: '#ffffff'
      };

      await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendarList.patch).toHaveBeenCalledWith({
        calendarId: 'cal@example.com',
        requestBody: {
          backgroundColor: '#0088aa',
          foregroundColor: '#ffffff'
        },
        colorRgbFormat: true
      });
    });

    it('calls both resources when both kinds of fields are provided', async () => {
      const args = {
        calendarId: 'cal@example.com',
        timeZone: 'America/New_York',
        colorId: '5'
      };

      const result = await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendars.patch).toHaveBeenCalledTimes(1);
      expect(mockCalendar.calendarList.patch).toHaveBeenCalledTimes(1);

      const response = JSON.parse(result.content[0].text);
      expect(response.updated).toEqual(['timeZone', 'colorId']);
    });

    it('forwards empty strings to clear shared fields (description/location)', async () => {
      // Empty string is a deliberate clear, not "no change" — guards use !== undefined
      const args = {
        calendarId: 'cal@example.com',
        description: '',
        location: ''
      };

      const result = await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendars.patch).toHaveBeenCalledWith({
        calendarId: 'cal@example.com',
        requestBody: { description: '', location: '' }
      });

      const response = JSON.parse(result.content[0].text);
      expect(response.updated).toEqual(['description', 'location']);
    });

    it('forwards an empty defaultReminders array to clear all reminders', async () => {
      const args = {
        calendarId: 'cal@example.com',
        defaultReminders: []
      };

      const result = await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendarList.patch).toHaveBeenCalledWith({
        calendarId: 'cal@example.com',
        requestBody: { defaultReminders: [] },
        colorRgbFormat: undefined
      });

      const response = JSON.parse(result.content[0].text);
      expect(response.updated).toEqual(['defaultReminders']);
    });

    it('passes defaultReminders and notificationSettings through to calendarList.patch', async () => {
      const args = {
        calendarId: 'cal@example.com',
        defaultReminders: [{ method: 'email', minutes: 30 }],
        notificationSettings: {
          notifications: [{ type: 'eventCreation', method: 'email' }]
        }
      };

      await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendarList.patch).toHaveBeenCalledWith({
        calendarId: 'cal@example.com',
        requestBody: {
          defaultReminders: [{ method: 'email', minutes: 30 }],
          notificationSettings: { notifications: [{ type: 'eventCreation', method: 'email' }] }
        },
        colorRgbFormat: undefined
      });
    });
  });

  describe('Response', () => {
    it('returns the re-read calendar and clears the registry cache', async () => {
      const clearSpy = vi.spyOn((handler as any).calendarRegistry, 'clearCache');

      const args = { calendarId: 'cal@example.com', colorId: '5' };
      const result = await handler.runTool(args, mockAccounts);

      expect(mockCalendar.calendarList.get).toHaveBeenCalledWith({ calendarId: 'cal@example.com' });
      expect(clearSpy).toHaveBeenCalled();

      const response = JSON.parse(result.content[0].text);
      expect(response.calendar.timeZone).toBe('America/New_York');
      expect(response.calendar.colorId).toBe('5');
    });

    it('resolves the account with write access', async () => {
      const spy = vi.spyOn(handler as any, 'getClientWithAutoSelection');
      const args = { calendarId: 'cal@example.com', colorId: '5', account: 'test' };

      await handler.runTool(args, mockAccounts);

      expect(spy).toHaveBeenCalledWith('test', 'cal@example.com', mockAccounts, 'write');
    });
  });

  describe('Error handling', () => {
    it('surfaces API errors via handleGoogleApiError', async () => {
      mockCalendar.calendars.patch.mockRejectedValue(Object.assign(new Error('Forbidden'), { code: 403 }));
      vi.spyOn(handler as any, 'handleGoogleApiError').mockImplementation(() => {
        throw new Error('Permission denied');
      });

      const args = { calendarId: 'cal@example.com', timeZone: 'America/New_York' };

      await expect(handler.runTool(args, mockAccounts)).rejects.toThrow('Permission denied');
    });
  });
});
