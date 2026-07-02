import { authConfig } from './config.js';
import { createAuthError } from '../utils/mcpErrorResponse.js';
import { isHeadlessMode } from './dataPaths.js';

export function shouldUseDeviceCodeFlow() {
  return isHeadlessMode();
}

export async function requestDeviceCode(clientId, authority) {
  const deviceCodeUrl = `https://login.microsoftonline.com/${authority}/oauth2/v2.0/devicecode`;
  const params = new URLSearchParams({
    client_id: clientId,
    scope: authConfig.oauth.scope,
  });

  const response = await fetch(deviceCodeUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });

  if (!response.ok) {
    const error = await response.text();
    throw createAuthError(`Device code request failed: ${error}`, true);
  }

  return await response.json();
}

export async function pollDeviceCodeToken(clientId, authority, deviceCode, interval = 5, expiresIn = 900) {
  const tokenUrl = `https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`;
  const deadline = Date.now() + expiresIn * 1000;
  let pollInterval = interval * 1000;

  while (Date.now() < deadline) {
    await new Promise(resolve => setTimeout(resolve, pollInterval));

    const params = new URLSearchParams({
      client_id: clientId,
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      device_code: deviceCode,
    });

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    const body = await response.json();

    if (response.ok) {
      return body;
    }

    if (body.error === 'authorization_pending') {
      console.error('Waiting for device code authentication...');
      continue;
    }

    if (body.error === 'slow_down') {
      pollInterval += 5000;
      continue;
    }

    throw createAuthError(`Device code token poll failed: ${body.error_description || body.error}`, true);
  }

  throw createAuthError('Device code authentication timed out', true);
}

export async function authenticateWithDeviceCode(clientId, authority) {
  const deviceCodeResponse = await requestDeviceCode(clientId, authority);
  console.error(deviceCodeResponse.message);

  const tokenResponse = await pollDeviceCodeToken(
    clientId,
    authority,
    deviceCodeResponse.device_code,
    deviceCodeResponse.interval,
    deviceCodeResponse.expires_in
  );

  return {
    tokenResponse,
    deviceCodeInfo: {
      status: 'pending',
      message: deviceCodeResponse.message,
      verificationUri: deviceCodeResponse.verification_uri,
      userCode: deviceCodeResponse.user_code,
      expiresIn: deviceCodeResponse.expires_in,
    },
  };
}
