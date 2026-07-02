import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { resolveReadAccount, fanOutAcrossAccounts } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

// List calendar events
export async function listEventsTool(registry, args) {
  const { startDateTime, endDateTime, limit = 10, calendar } = args;

  try {
    const buildEventResponse = async (manager, scopedArgs) => {
      await manager.ensureAuthenticated();
      const graphApiClient = manager.getGraphApiClient();
      const mailboxBase = buildMailboxBase(scopedArgs.mailbox);

      // Use calendarView endpoint which properly handles timezones and recurring events.
      const endpoint = scopedArgs.calendar
        ? `${mailboxBase}/calendars/${scopedArgs.calendar}/calendarView`
        : `${mailboxBase}/calendarView`;

      const options = {
        select: 'id,subject,start,end,location,attendees,bodyPreview,organizer,isAllDay,showAs,importance,sensitivity,categories,webLink',
        top: scopedArgs.limit ?? limit,
        orderby: 'start/dateTime',
      };

      if (scopedArgs.startDateTime && scopedArgs.endDateTime) {
        options.startDateTime = scopedArgs.startDateTime;
        options.endDateTime = scopedArgs.endDateTime;
      } else {
        const today = new Date();
        const tomorrow = new Date(today);
        tomorrow.setDate(tomorrow.getDate() + 1);
        options.startDateTime = today.toISOString();
        options.endDateTime = tomorrow.toISOString();
      }

      const result = await graphApiClient.makeRequest(endpoint, options);
      const events = (result.value || []).map(event => ({
        id: event.id,
        subject: event.subject,
        start: event.start,
        end: event.end,
        location: event.location?.displayName || 'No location',
        attendees: event.attendees?.map(a => a.emailAddress?.address) || [],
        preview: event.bodyPreview,
        organizer: event.organizer?.emailAddress?.address || 'Unknown',
        isAllDay: event.isAllDay,
        webLink: event.webLink,
      }));
      return { events };
    };

    if (!args.account) {
      const response = await fanOutAcrossAccounts(
        registry,
        args,
        (manager, account, scopedArgs) => buildEventResponse(manager, scopedArgs),
        {
          resultKey: 'events',
          sortKey: 'start.dateTime',
          globalLimit: limit,
        }
      );
      return createSafeResponse({
        ...response,
        count: response.events.length,
      });
    }

    const { manager } = await resolveReadAccount(registry, args);
    const { events } = await buildEventResponse(manager, args);
    return createSafeResponse({ events, count: events.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list events');
  }
}