#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ListPromptsRequestSchema,
  GetPromptRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  console.error('Stack trace:', reason.stack || 'No stack trace available');
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  console.error('Stack trace:', error.stack);
  if (error.code === 'MODULE_NOT_FOUND' || error.name === 'SyntaxError') {
    process.exit(1);
  }
});

let catalogContextPromise = null;
let runtimeContextPromise = null;

async function getCatalogContext() {
  if (!catalogContextPromise) {
    catalogContextPromise = loadCatalogContext();
  }
  return catalogContextPromise;
}

async function loadCatalogContext() {
  await import('dotenv/config');
  const { allToolSchemas } = await import('./schemas/toolSchemas.js');
  const { promptList, getPrompt } = await import('./prompts/index.js');
  return { allToolSchemas, promptList, getPrompt };
}

// Preload catalog before the client sends tools/list (must not compete with runtime imports).
void getCatalogContext();

async function getRuntimeContext() {
  if (!runtimeContextPromise) {
    runtimeContextPromise = loadRuntimeContext();
  }
  return runtimeContextPromise;
}

async function loadRuntimeContext() {
  console.error('Debug: Loading application modules...');
  await import('dotenv/config');

  const { authManagerRegistry } = await import('./auth/authManagerRegistry.js');
  const { getStartupConfig } = await import('./auth/defaultApp.js');
  const { createProtocolError, ErrorCodes, convertErrorToToolError } = await import('./utils/mcpErrorResponse.js');
  const tools = await import('./tools/index.js');

  const registry = authManagerRegistry;
  await registry.initialize();

  const startupConfig = getStartupConfig();
  console.error(`Debug: Auth mode = ${startupConfig.authMode}`);
  console.error(`Debug: AZURE_CLIENT_ID = ${process.env.AZURE_CLIENT_ID ? 'SET (BYO)' : 'using default'}`);
  console.error(`Debug: AZURE_TENANT_ID = ${process.env.AZURE_TENANT_ID ? 'SET (BYO)' : 'using organizations'}`);
  console.error('Call outlook_connect_account to sign in, or use an existing connected account.');

  return {
    registry,
    tools,
    createProtocolError,
    ErrorCodes,
    convertErrorToToolError,
  };
}

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '1.0.2',
  },
  {
    capabilities: {
      tools: {},
      prompts: {},
    },
  }
);

server.oninitialized = () => {
  console.error('Debug: Client initialized');
};

server.setRequestHandler(ListToolsRequestSchema, async () => {
  const { allToolSchemas } = await getCatalogContext();
  console.error(`Debug: Returning ${allToolSchemas.length} tools to client`);
  return { tools: allToolSchemas };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const ctx = await getRuntimeContext();
  const { name, arguments: args } = request.params;
  console.error(`DEBUG Tool Dispatch: Called tool '${name}' with args:`, JSON.stringify(args, null, 2));

  const {
    registry,
    tools,
    createProtocolError,
    ErrorCodes,
    convertErrorToToolError,
  } = ctx;

  try {
    switch (name) {
      case 'outlook_list_emails':
        return await tools.listEmailsTool(registry, args);
      case 'outlook_send_email':
        return await tools.sendEmailTool(registry, args);
      case 'outlook_list_events':
        return await tools.listEventsTool(registry, args);
      case 'outlook_create_event':
        return await tools.createEventTool(registry, args);
      case 'outlook_get_event':
        return await tools.getEventTool(registry, args);
      case 'outlook_update_event':
        return await tools.updateEventTool(registry, args);
      case 'outlook_delete_event':
        return await tools.deleteEventTool(registry, args);
      case 'outlook_respond_to_invite':
        return await tools.respondToInviteTool(registry, args);
      case 'outlook_validate_event_datetimes':
        return await tools.validateEventDateTimesTool(registry, args);
      case 'outlook_create_recurring_event':
        return await tools.createRecurringEventTool(registry, args);
      case 'outlook_find_meeting_times':
        return await tools.findMeetingTimesTool(registry, args);
      case 'outlook_check_availability':
        return await tools.checkAvailabilityTool(registry, args);
      case 'outlook_schedule_online_meeting':
        return await tools.scheduleOnlineMeetingTool(registry, args);
      case 'outlook_list_calendars':
        return await tools.listCalendarsTool(registry, args);
      case 'outlook_get_calendar_view':
        return await tools.getCalendarViewTool(registry, args);
      case 'outlook_get_busy_times':
        return await tools.getBusyTimesTool(registry, args);
      case 'outlook_build_recurrence_pattern':
        return await tools.buildRecurrencePatternTool(registry, args);
      case 'outlook_create_recurrence_helper':
        return await tools.createRecurrenceHelperTool(registry, args);
      case 'outlook_check_calendar_permissions':
        return await tools.checkCalendarPermissionsTool(registry, args);
      case 'outlook_get_email':
        return await tools.getEmailTool(registry, args);
      case 'outlook_search_emails':
        return await tools.searchEmailsTool(registry, args);
      case 'outlook_create_draft':
        return await tools.createDraftTool(registry, args);
      case 'outlook_reply_to_email':
        return await tools.replyToEmailTool(registry, args);
      case 'outlook_reply_all':
        return await tools.replyAllTool(registry, args);
      case 'outlook_forward_email':
        return await tools.forwardEmailTool(registry, args);
      case 'outlook_delete_email':
        return await tools.deleteEmailTool(registry, args);
      case 'outlook_move_email':
        return await tools.moveEmailTool(registry, args);
      case 'outlook_mark_as_read':
        return await tools.markAsReadTool(registry, args);
      case 'outlook_flag_email':
        return await tools.flagEmailTool(registry, args);
      case 'outlook_categorize_email':
        return await tools.categorizeEmailTool(registry, args);
      case 'outlook_archive_email':
        return await tools.archiveEmailTool(registry, args);
      case 'outlook_batch_process_emails':
        return await tools.batchProcessEmailsTool(registry, args);
      case 'outlook_list_folders':
        return await tools.listFoldersTool(registry, args);
      case 'outlook_create_folder':
        return await tools.createFolderTool(registry, args);
      case 'outlook_rename_folder':
        return await tools.renameFolderTool(registry, args);
      case 'outlook_get_folder_stats':
        return await tools.getFolderStatsTool(registry, args);
      case 'outlook_list_attachments':
        return await tools.listAttachmentsTool(registry, args);
      case 'outlook_download_attachment':
        return await tools.downloadAttachmentTool(registry, args);
      case 'outlook_add_attachment':
        return await tools.addAttachmentTool(registry, args);
      case 'outlook_scan_attachments':
        return await tools.scanAttachmentsTool(registry, args);
      case 'outlook_get_sharepoint_file':
        return await tools.getSharePointFileTool(registry, args);
      case 'outlook_list_sharepoint_files':
        return await tools.listSharePointFilesTool(registry, args);
      case 'outlook_resolve_sharepoint_link':
        return await tools.resolveSharePointLinkTool(registry, args);
      case 'outlook_connect_account':
        return await tools.connectAccountTool(registry, args);
      case 'outlook_list_accounts':
        return await tools.listAccountsTool(registry, args);
      case 'outlook_disconnect_account':
        return await tools.disconnectAccountTool(registry, args);
      case 'outlook_set_default_account':
        return await tools.setDefaultAccountTool(registry, args);
      case 'outlook_list_accessible_mailboxes':
        return await tools.listAccessibleMailboxesTool(registry, args);
      default:
        return createProtocolError(
          ErrorCodes.METHOD_NOT_FOUND,
          `Unknown tool: ${name}`,
          { availableTools: Object.keys(tools).filter(key => key.endsWith('Tool')) }
        );
    }
  } catch (error) {
    console.error('Unexpected error in tool handler:', error);
    if (error.content && error.isError !== undefined) {
      return error;
    }
    return convertErrorToToolError(error, 'Tool execution failed');
  }
});

server.setRequestHandler(ListPromptsRequestSchema, async () => {
  const { promptList } = await getCatalogContext();
  console.error(`Debug: Returning ${promptList.length} prompts to client`);
  return { prompts: promptList };
});

server.setRequestHandler(GetPromptRequestSchema, async (request) => {
  const { getPrompt } = await getCatalogContext();
  const { name, arguments: args } = request.params;
  console.error(`Debug: Getting prompt '${name}' with args:`, JSON.stringify(args, null, 2));
  try {
    return await getPrompt(name, args);
  } catch (error) {
    console.error(`Error getting prompt ${name}:`, error);
    const { convertErrorToToolError } = await import('./utils/mcpErrorResponse.js');
    throw convertErrorToToolError(error, `Failed to get prompt ${name}`);
  }
});

try {
  console.error('Debug: Connecting MCP transport...');
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP server is ready and connected');
} catch (error) {
  console.error('Server error:', error);
  process.exit(1);
}
