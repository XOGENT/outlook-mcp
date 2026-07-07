import { applyUserStyling, clearStylingCache } from '../common/sharedUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';
import { resolveWriteAccount } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';
import { withSendDedupe, buildSendKey } from '../common/sendDedupe.js';

// Send email with user styling
export async function sendEmailTool(registry, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [], preserveUserStyling = true, idempotencyKey } = args;

  try {
    const writeResolution = await resolveWriteAccount(registry, args);
    if (writeResolution?.isError) {
      return writeResolution;
    }
    const { manager, account } = writeResolution;
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const mailboxBase = buildMailboxBase(args.mailbox);

    const dedupeKey = buildSendKey(
      idempotencyKey
        ? { accountId: account.accountId, idempotencyKey }
        : { accountId: account.accountId, to, cc, bcc, subject, body, bodyType, mailbox: args.mailbox || null }
    );

    let finalBody = body;
    let finalBodyType = bodyType;

    // If preserving user styling, get user's default styling and signature
    if (preserveUserStyling) {
      const styledBody = await applyUserStyling(graphApiClient, body, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    const message = {
      subject,
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
    };

    if (cc.length > 0) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    // Guard against a duplicate delivery when an identical send is re-issued
    // after an ambiguous failure (transport timeout then a model retry).
    return await withSendDedupe(dedupeKey, async () => {
      // makeRequest RETURNS handled errors (it does not throw), so an ambiguous
      // outcome or a hard failure must be propagated — otherwise we'd falsely
      // report success and the caller might resend a message that already went.
      const sendResult = await graphApiClient.postWithRetry(`${mailboxBase}/sendMail`, {
        message,
        saveToSentItems: true,
      });
      if (sendResult?.isError) {
        return sendResult;
      }

      // Invalidate styling cache after sending email (user might have changed styling)
      // Don't invalidate signature cache as frequently since signatures change less often
      try {
        const userInfo = await graphApiClient.makeRequest('/me', { select: 'id' });
        clearStylingCache(userInfo.id);
      } catch (error) {
        console.warn('Could not invalidate styling cache:', error.message);
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              message: `Email sent successfully to ${to.join(', ')}`,
              fromAccount: {
                accountId: account.accountId,
                email: account.email
              }
            }, null, 2),
          },
        ],
      };
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to send email');
  }
}