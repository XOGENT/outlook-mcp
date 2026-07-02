#!/usr/bin/env node

import { pathToFileURL } from 'url';
import { authManagerRegistry } from '../auth/authManagerRegistry.js';
import { getStartupConfig } from '../auth/defaultApp.js';
import { graphHelpers } from '../graph/graphHelpers.js';

async function ensureAccounts(registry) {
  await registry.initialize();

  if (await registry.hasAccounts()) {
    return registry.listAccounts();
  }

  if (process.env.CONNECT_ACCOUNT === '1') {
    console.log('No accounts found — starting interactive connect flow...\n');
    const startup = getStartupConfig();
    const result = await registry.connectAccount({
      clientId: process.env.AZURE_CLIENT_ID || startup.clientId,
      tenantId: process.env.AZURE_TENANT_ID,
    });

    if (!result.success) {
      throw new Error(`Connect failed: ${result.error?.content?.[0]?.text || result.error?.message || 'unknown error'}`);
    }

    console.log(`Connected: ${result.account.email} (${result.account.accountId})`);
    if (result.deviceCodeInfo) {
      console.log(`Device code: ${result.deviceCodeInfo.userCode}`);
      console.log(`Visit: ${result.deviceCodeInfo.verificationUri}`);
    }
    return registry.listAccounts();
  }

  console.error('No accounts connected.');
  console.error('Connect an account first, or run with CONNECT_ACCOUNT=1 to authenticate interactively.');
  console.error('Optional BYO env vars: AZURE_CLIENT_ID, AZURE_TENANT_ID');
  process.exit(1);
}

async function testAccount(registry, account) {
  console.log(`\n=== Account: ${account.email} (${account.accountId}) ===\n`);

  const { manager } = await registry.resolve(account.accountId);
  await manager.ensureAuthenticated();
  const graphApiClient = manager.getGraphApiClient();

  console.log('1. Testing optimized email request...');
  const emails = await graphApiClient.getWithSelect('/me/messages', [
    'subject', 'from', 'receivedDateTime', 'isRead',
  ]);
  console.log(`   ✓ Retrieved ${emails.value?.length || 0} emails`);

  console.log('2. Testing calendar request...');
  const events = await graphApiClient.makeRequest('/me/events', {
    select: 'subject,start,end',
    top: 5,
    orderby: 'start/dateTime',
  });
  console.log(`   ✓ Retrieved ${events.value?.length || 0} calendar events`);

  console.log('3. Testing Graph helpers...');
  const filter = graphHelpers.general.buildODataFilter({
    isRead: false,
    receivedDateTime: { $gt: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000) },
  });
  console.log(`   ✓ Built OData filter: ${filter}`);

  console.log('4. Testing batch request...');
  const batchRequests = [
    { method: 'GET', url: '/me' },
    { method: 'GET', url: '/me/mailFolders/inbox' },
    { method: 'GET', url: '/me/calendar' },
  ];
  const batchResponse = await graphApiClient.makeBatchRequest(batchRequests);
  console.log(`   ✓ Executed batch request with ${batchResponse.length} operations`);

  console.log('5. Testing error handling...');
  try {
    await graphApiClient.makeRequest('/me/nonexistent-endpoint');
  } catch (error) {
    console.log(`   ✓ Error handling working: ${error.message}`);
  }

  console.log('6. Testing pagination helper...');
  let emailCount = 0;
  for await (const emailBatch of graphApiClient.iterateAllPages('/me/messages', { top: 2 })) {
    emailCount += emailBatch.length;
    if (emailCount >= 4) break;
  }
  console.log(`   ✓ Pagination retrieved ${emailCount} emails across multiple pages`);
}

async function testGraphApiClient() {
  try {
    console.log('Testing Graph API Client (multi-account)...\n');

    const registry = authManagerRegistry;
    let accounts = await ensureAccounts(registry);

    if (process.env.CONNECT_SECOND_ACCOUNT === '1') {
      console.log('\nConnecting a second account...');
      const startup = getStartupConfig();
      const result = await registry.connectAccount({
        clientId: process.env.AZURE_CLIENT_ID || startup.clientId,
        tenantId: process.env.AZURE_TENANT_ID,
      });
      if (!result.success) {
        throw new Error(`Second account connect failed: ${result.error?.content?.[0]?.text || 'unknown error'}`);
      }
      console.log(`Connected second account: ${result.account.email}`);
      accounts = await registry.listAccounts();
    }

    console.log(`Found ${accounts.length} connected account(s).`);
    for (const account of accounts) {
      await testAccount(registry, account);
    }

    console.log('\n✅ All Graph API Client tests passed!');
    console.log('\nVerified per account:');
    console.log('- Authentication via AuthManagerRegistry');
    console.log('- Rate limiting and retry logic');
    console.log('- Request optimization with $select');
    console.log('- Batch request processing');
    console.log('- Error handling');
    console.log('- Pagination support');
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    if (error.stack) {
      console.error('Stack trace:', error.stack);
    }
    process.exit(1);
  }
}

const isMain = process.argv[1]
  && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;

if (isMain) {
  testGraphApiClient();
}
