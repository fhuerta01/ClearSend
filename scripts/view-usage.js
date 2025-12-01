#!/usr/bin/env node

/**
 * View ClearSend Usage Statistics
 *
 * This script queries the Vercel KV database to view the current usage count.
 *
 * Requirements:
 * - @vercel/kv package installed
 * - Vercel KV database created and configured
 * - Environment variables set (via `vercel env pull`)
 *
 * Usage:
 *   node scripts/view-usage.js
 *
 * Or add to package.json:
 *   "scripts": {
 *     "usage": "node scripts/view-usage.js"
 *   }
 *
 * Then run: npm run usage
 */

async function viewUsageCount() {
  try {
    // Import Vercel KV
    const { kv } = await import('@vercel/kv');

    const USAGE_KEY = 'clearsend:usage:count';

    console.log('📊 Fetching ClearSend usage statistics...\n');

    // Get the count
    const count = await kv.get(USAGE_KEY);

    if (count === null) {
      console.log('❌ No usage data found.');
      console.log('   This could mean:');
      console.log('   1. Analytics haven\'t been deployed yet');
      console.log('   2. Nobody has used the add-in yet');
      console.log('   3. Wrong Vercel KV database connected\n');
      return;
    }

    console.log('✅ ClearSend Usage Count');
    console.log('─────────────────────────');
    console.log(`   Total loads: ${count.toLocaleString()}`);
    console.log('─────────────────────────\n');

    console.log('📝 Note: This is just a counter of how many times');
    console.log('   the add-in was loaded. No personal data is collected.\n');

  } catch (error) {
    console.error('❌ Error fetching usage count:');

    if (error.message.includes('KV_REST_API_URL')) {
      console.error('\n   Missing Vercel KV environment variables.');
      console.error('   Run this command first:');
      console.error('   → vercel env pull\n');
    } else if (error.code === 'MODULE_NOT_FOUND') {
      console.error('\n   @vercel/kv package not installed.');
      console.error('   Install it first:');
      console.error('   → npm install @vercel/kv\n');
    } else {
      console.error(`   ${error.message}\n`);
    }

    process.exit(1);
  }
}

// Run the script
viewUsageCount();
