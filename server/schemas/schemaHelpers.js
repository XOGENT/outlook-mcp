export const readAccountParam = {
  type: 'string',
  description: 'Account ID (from outlook_list_accounts). Defaults to the default account.',
};

export const fanOutAccountParam = {
  type: 'string',
  description: 'Optional. Restrict to one account. Omit to query ALL connected accounts.',
};

export const writeAccountParam = {
  type: 'string',
  description: 'Account ID to send from (from outlook_list_accounts). Required when multiple accounts are connected.',
};

export const mailboxParam = {
  type: 'string',
  description: 'Shared/delegated mailbox email or user ID. Omit for the signed-in user primary mailbox.',
};

const FAN_OUT_TOOLS = new Set([
  'outlook_search_emails',
  'outlook_list_events',
  'outlook_get_calendar_view',
  'outlook_list_calendars',
  'outlook_get_busy_times',
]);

const WRITE_TOOLS = new Set([
  'outlook_send_email',
  'outlook_reply_to_email',
  'outlook_reply_all',
  'outlook_forward_email',
  'outlook_create_draft',
  'outlook_create_event',
  'outlook_update_event',
  'outlook_delete_event',
  'outlook_respond_to_invite',
  'outlook_create_recurring_event',
  'outlook_schedule_online_meeting',
  'outlook_create_recurrence_helper',
  'outlook_delete_email',
  'outlook_move_email',
  'outlook_mark_as_read',
  'outlook_flag_email',
  'outlook_categorize_email',
  'outlook_archive_email',
  'outlook_batch_process_emails',
  'outlook_create_folder',
  'outlook_rename_folder',
  'outlook_add_attachment',
]);

const ACCOUNT_TOOLS = new Set([
  'outlook_connect_account',
  'outlook_list_accounts',
  'outlook_disconnect_account',
  'outlook_set_default_account',
  'outlook_list_accessible_mailboxes',
]);

export function augmentSchema(schema) {
  if (ACCOUNT_TOOLS.has(schema.name)) return schema;

  const properties = { ...(schema.inputSchema?.properties || {}) };

  if (!properties.account) {
    if (FAN_OUT_TOOLS.has(schema.name)) {
      properties.account = fanOutAccountParam;
    } else if (WRITE_TOOLS.has(schema.name)) {
      properties.account = writeAccountParam;
    } else {
      properties.account = readAccountParam;
    }
  }

  if (!properties.mailbox && schema.name !== 'outlook_validate_event_datetimes'
    && schema.name !== 'outlook_build_recurrence_pattern') {
    properties.mailbox = mailboxParam;
  }

  return {
    ...schema,
    inputSchema: {
      ...schema.inputSchema,
      properties,
    },
  };
}

export function augmentSchemas(schemas) {
  return schemas.map(augmentSchema);
}
