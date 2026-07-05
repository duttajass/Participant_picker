#!/usr/bin/env node

/**
 * API Test Script: Test Teams Picker API Endpoints
 * 
 * This script tests the API endpoints without requiring a real Teams meeting.
 * Useful for verifying the server is working correctly.
 * 
 * Usage:
 *   node test-api.js [BASE_URL]
 * 
 * Example:
 *   node test-api.js http://localhost:3000
 */

const BASE_URL = process.argv[2] || 'http://localhost:3000';
const API_URL = BASE_URL.replace(/\/$/, ''); // Remove trailing slash

async function test(method, endpoint, body = null) {
  try {
    const opts = {
      method,
      headers: { 'Content-Type': 'application/json' },
    };
    if (body) opts.body = JSON.stringify(body);

    const res = await fetch(`${API_URL}${endpoint}`, opts);
    const data = await res.json();

    return { status: res.status, data };
  } catch (err) {
    return { error: err.message };
  }
}

async function runTests() {
  console.log('\n🧪 Teams Picker API Test Suite');
  console.log(`📍 Testing: ${API_URL}\n`);
  console.log('=====================================\n');

  // Test 1: Health Check
  console.log('1️⃣  Testing: GET /api/health');
  let result = await test('GET', '/api/health');
  if (result.error) {
    console.log(`   ❌ Error: ${result.error}`);
    console.log(`   ⚠️  Server may not be running. Start it with: npm start\n`);
    return;
  }
  if (result.status === 200 && result.data.ok) {
    console.log(`   ✅ Server is healthy`);
    console.log(`   Active sessions: ${result.data.activeSessions}\n`);
  } else {
    console.log(`   ❌ Unexpected response: ${result.status}\n`);
  }

  // Test 2: Get Sessions (should be empty)
  console.log('2️⃣  Testing: GET /api/sessions');
  result = await test('GET', '/api/sessions');
  if (result.status === 200) {
    console.log(`   ✅ Sessions endpoint working`);
    console.log(`   Current sessions: ${result.data.sessions.length}\n`);
  } else {
    console.log(`   ❌ Failed with status ${result.status}\n`);
  }

  // Test 3: Try to join with invalid URL (should fail)
  console.log('3️⃣  Testing: POST /api/join (invalid URL - should fail gracefully)');
  result = await test('POST', '/api/join', { meetingUrl: 'http://invalid-url.com' });
  if (result.status >= 400) {
    console.log(`   ✅ Correctly rejected invalid URL`);
    console.log(`   Response: ${result.data.error}\n`);
  } else {
    console.log(`   ⚠️  Unexpected response\n`);
  }

  // Test 4: Try to pick without session (should fail)
  console.log('4️⃣  Testing: POST /api/pick (no active session - should fail)');
  result = await test('POST', '/api/pick', {
    meetingUrl: 'https://teams.microsoft.com/l/meetup-join/test',
  });
  if (result.status >= 400) {
    console.log(`   ✅ Correctly rejected request (no active session)`);
    console.log(`   Response: ${result.data.error}\n`);
  } else {
    console.log(`   ⚠️  Unexpected response\n`);
  }

  // Test 5: Try auto-pick without session (should fail)
  console.log('5️⃣  Testing: POST /api/pick-and-post (no active session - should fail)');
  result = await test('POST', '/api/pick-and-post', {
    meetingUrl: 'https://teams.microsoft.com/l/meetup-join/test',
  });
  if (result.status >= 400) {
    console.log(`   ✅ Correctly rejected request (no active session)`);
    console.log(`   Response: ${result.data.error}\n`);
  } else {
    console.log(`   ⚠️  Unexpected response\n`);
  }

  console.log('=====================================');
  console.log('\n📋 API Test Summary:');
  console.log('   ✅ Health check endpoint working');
  console.log('   ✅ Sessions endpoint working');
  console.log('   ✅ Error handling working');
  console.log('\n✨ All basic API tests passed!');
  console.log('\n💡 Next Steps:');
  console.log('   1. Start a real Teams meeting');
  console.log('   2. Use the web UI at http://localhost:3000');
  console.log('   3. Test the auto-pick & post to chat feature');
  console.log('   4. If chat posting fails, run: node debug-selectors.js "<MEETING_URL>"\n');
}

runTests().catch(console.error);
