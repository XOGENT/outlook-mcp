import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { resolveWriteAccount } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

// Forward an email
export async function forwardEmailTool(registry, args) {
  const { messageId, to, body = '', bodyType = 'text', comment = '', preserveUserStyling = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!to || to.length === 0) {
    return createValidationError('to', 'At least one recipient is required');
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

    const forwardPayload = {
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
    };

    // Use body or comment as the forward message text
    const forwardText = body || comment;
    if (forwardText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, forwardText, bodyType);
        // For forward API, we need to strip HTML tags and use plain text in comment
        forwardPayload.comment = styledBody.type === 'html' ? 
          styledBody.content.replace(/<[^>]*>/g, '') : 
          styledBody.content;
      } else {
        forwardPayload.comment = forwardText;
      }
    }

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${messageId}/forward`, forwardPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Email forwarded successfully to ${to.join(', ')}. Forward ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to forward email');
  }
}