import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { resolveWriteAccount } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

// Reply to an email
export async function replyToEmailTool(registry, args) {
  const { messageId, body, bodyType = 'text', comment = '', preserveUserStyling = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!body && !comment) {
    return createValidationError('body/comment', 'Either body or comment is required');
  }

  try {
    const writeResolution = await resolveWriteAccount(registry, args);
    if (writeResolution?.isError) {
      return writeResolution;
    }
    const { manager } = writeResolution;
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const mailboxBase = buildMailboxBase(args.mailbox);

    const replyPayload = {};

    // Use body or comment as the reply message text
    const replyText = body || comment;
    if (replyText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, replyText, bodyType);
        replyPayload.message = {
          body: {
            contentType: styledBody.type === 'html' ? 'HTML' : 'Text',
            content: styledBody.content,
          },
        };
      } else {
        replyPayload.message = {
          body: {
            contentType: bodyType === 'html' ? 'HTML' : 'Text',
            content: replyText,
          },
        };
      }
    }

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${messageId}/reply`, replyPayload);
    if (result?.isError) {
      return result;
    }

    return {
      content: [
        {
          type: 'text',
          text: `Reply created successfully. Reply ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reply to email');
  }
}

// Reply all to an email
export async function replyAllTool(registry, args) {
  const { messageId, body, bodyType = 'text', comment = '', preserveUserStyling = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!body && !comment) {
    return createValidationError('body/comment', 'Either body or comment is required');
  }

  try {
    const writeResolution = await resolveWriteAccount(registry, args);
    if (writeResolution?.isError) {
      return writeResolution;
    }
    const { manager } = writeResolution;
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const mailboxBase = buildMailboxBase(args.mailbox);

    const replyPayload = {};

    // Use body or comment as the reply message text
    const replyText = body || comment;
    if (replyText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, replyText, bodyType);
        replyPayload.message = {
          body: {
            contentType: styledBody.type === 'html' ? 'HTML' : 'Text',
            content: styledBody.content,
          },
        };
      } else {
        replyPayload.message = {
          body: {
            contentType: bodyType === 'html' ? 'HTML' : 'Text',
            content: replyText,
          },
        };
      }
    }

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${messageId}/replyAll`, replyPayload);
    if (result?.isError) {
      return result;
    }

    return {
      content: [
        {
          type: 'text',
          text: `Reply all created successfully. Reply ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reply all to email');
  }
}