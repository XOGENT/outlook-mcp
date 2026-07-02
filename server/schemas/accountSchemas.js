/**
 * Account-related MCP tool schemas
 */

export const connectAccountSchema = {
  name: 'outlook_connect_account',
  description: 'Connect a Microsoft account via OAuth. Opens browser (desktop) or returns device code (headless). Optional BYO Azure app credentials.',
  inputSchema: {
    type: 'object',
    properties: {
      clientId: {
        type: 'string',
        description: 'Optional BYO Azure application client ID',
      },
      tenantId: {
        type: 'string',
        description: 'Optional BYO Azure tenant ID. Omit for multi-tenant hosted app.',
      },
    },
  },
};

export const listAccountsSchema = {
  name: 'outlook_list_accounts',
  description: 'List all connected Microsoft accounts',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const disconnectAccountSchema = {
  name: 'outlook_disconnect_account',
  description: 'Disconnect and remove a connected account',
  inputSchema: {
    type: 'object',
    properties: {
      accountId: {
        type: 'string',
        description: 'Account ID to disconnect',
      },
    },
    required: ['accountId'],
  },
};

export const setDefaultAccountSchema = {
  name: 'outlook_set_default_account',
  description: 'Set the default account for single-account tools',
  inputSchema: {
    type: 'object',
    properties: {
      accountId: {
        type: 'string',
        description: 'Account ID to set as default',
      },
    },
    required: ['accountId'],
  },
};

export const listAccessibleMailboxesSchema = {
  name: 'outlook_list_accessible_mailboxes',
  description: 'List shared and delegated mailboxes accessible to an account',
  inputSchema: {
    type: 'object',
    properties: {
      accountId: {
        type: 'string',
        description: 'Account ID (defaults to default account)',
      },
    },
  },
};

export const accountSchemas = [
  connectAccountSchema,
  listAccountsSchema,
  disconnectAccountSchema,
  setDefaultAccountSchema,
  listAccessibleMailboxesSchema,
];

export const accountSchemaMap = {
  outlook_connect_account: connectAccountSchema,
  outlook_list_accounts: listAccountsSchema,
  outlook_disconnect_account: disconnectAccountSchema,
  outlook_set_default_account: setDefaultAccountSchema,
  outlook_list_accessible_mailboxes: listAccessibleMailboxesSchema,
};
