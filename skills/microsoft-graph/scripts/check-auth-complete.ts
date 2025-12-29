#!/usr/bin/env bun
/**
 * Microsoft Graph Auth Completion Check
 *
 * Verifies that authentication has completed after device code flow.
 * Call this after user says they've authenticated.
 *
 * Usage:
 *   bun run check-auth-complete.ts [--profile <name>] [--json]
 *
 * Output (JSON mode):
 *   { "status": "complete" | "pending" | "failed", "account"?: string, "error"?: string }
 *
 * Exit codes:
 *   0 - Authentication complete
 *   1 - Still pending or failed
 */

import { parseArgs } from "util";
import { getCredential, isTokenExpired } from "./lib/credentials";

const { values } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    profile: { type: "string", default: "default" },
    json: { type: "boolean", default: false },
    help: { type: "boolean", short: "h", default: false },
  },
});

interface CheckResult {
  status: "complete" | "pending" | "failed";
  account?: string;
  error?: string;
}

function output(result: CheckResult): void {
  if (values.json) {
    console.log(JSON.stringify(result));
  } else {
    if (result.status === "complete") {
      console.log(`Authentication complete!`);
      if (result.account) console.log(`Account: ${result.account}`);
    } else if (result.status === "pending") {
      console.log(`Authentication still pending`);
      if (result.error) console.log(`Note: ${result.error}`);
    } else {
      console.log(`Authentication failed`);
      if (result.error) console.log(`Error: ${result.error}`);
    }
  }
  process.exit(result.status === "complete" ? 0 : 1);
}

if (values.help) {
  console.log(`
Microsoft Graph Auth Completion Check

Usage:
  bun run check-auth-complete.ts [options]

Options:
  --profile <name>    Credential profile name (default: "default")
  --json              Output result as JSON
  -h, --help          Show this help message

Output (JSON mode):
  { "status": "complete" | "pending" | "failed", "account"?: string, "error"?: string }

Exit codes:
  0 - Authentication complete
  1 - Still pending or failed
`);
  process.exit(0);
}

async function main() {
  const profile = values.profile!;

  // Load credential
  const credential = await getCredential("microsoft-graph", profile);

  if (!credential) {
    output({ status: "pending", error: "No credentials found yet" });
    return;
  }

  // Check if we have a valid (non-expired) token
  if (!isTokenExpired(credential)) {
    output({
      status: "complete",
      account: credential.account,
    });
    return;
  }

  // Token exists but is expired - auth may have failed or user used old credentials
  output({
    status: "failed",
    error: "Credentials expired - please try authenticating again",
  });
}

main();
