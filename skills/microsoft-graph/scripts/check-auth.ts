#!/usr/bin/env bun
/**
 * Microsoft Graph Auth Status Check
 *
 * Checks if authentication is valid, attempting silent refresh if needed.
 * Does not require user interaction - just reports status.
 *
 * Usage:
 *   bun run check-auth.ts [--profile <name>] [--json]
 *
 * Output (JSON mode):
 *   { "status": "valid" | "needs-auth", "expiresIn"?: number, "account"?: string, "error"?: string }
 *
 * Exit codes:
 *   0 - Valid token available
 *   1 - Authentication required
 */

import { parseArgs } from "util";
import {
  getCredential,
  setCredential,
  isTokenExpired,
} from "./lib/credentials";
import { PublicClientApplication } from "@azure/msal-node";
import { GRAPH_SCOPES } from "./lib/graph-client";
import type { Credential } from "./lib/types";

const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const DEFAULT_TENANT = "common";

const { values } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    profile: { type: "string", default: "default" },
    json: { type: "boolean", default: false },
    help: { type: "boolean", short: "h", default: false },
  },
});

interface CheckResult {
  status: "valid" | "needs-auth";
  expiresIn?: number;
  account?: string;
  error?: string;
}

function output(result: CheckResult): void {
  if (values.json) {
    console.log(JSON.stringify(result));
  } else {
    if (result.status === "valid") {
      console.log(`Auth status: Valid`);
      if (result.account) console.log(`Account: ${result.account}`);
      if (result.expiresIn !== undefined) {
        const mins = Math.round(result.expiresIn / 60);
        console.log(`Expires in: ${mins} minutes`);
      }
    } else {
      console.log(`Auth status: Authentication required`);
      if (result.error) console.log(`Reason: ${result.error}`);
    }
  }
  process.exit(result.status === "valid" ? 0 : 1);
}

if (values.help) {
  console.log(`
Microsoft Graph Auth Status Check

Usage:
  bun run check-auth.ts [options]

Options:
  --profile <name>    Credential profile name (default: "default")
  --json              Output result as JSON
  -h, --help          Show this help message

Output (JSON mode):
  { "status": "valid" | "needs-auth", "expiresIn"?: number, "account"?: string, "error"?: string }

Exit codes:
  0 - Valid token available
  1 - Authentication required (run auth.ts)
`);
  process.exit(0);
}

async function main() {
  const profile = values.profile!;

  // Load credential
  const credential = await getCredential("microsoft-graph", profile);

  if (!credential) {
    output({ status: "needs-auth", error: "No credentials found" });
    return;
  }

  // Check if token is valid (not expired)
  if (!isTokenExpired(credential)) {
    const expiresAt = new Date(credential.expiresAt);
    const expiresIn = Math.max(0, Math.floor((expiresAt.getTime() - Date.now()) / 1000));
    output({
      status: "valid",
      expiresIn,
      account: credential.account,
    });
    return;
  }

  // Token expired - attempt silent refresh
  if (!credential.refreshToken) {
    output({ status: "needs-auth", error: "Token expired, no refresh token" });
    return;
  }

  try {
    const clientId = credential.clientId ?? DEFAULT_CLIENT_ID;
    const tenantId = credential.tenantId ?? DEFAULT_TENANT;

    const pca = new PublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    });

    const scopes = [
      ...GRAPH_SCOPES.user,
      ...GRAPH_SCOPES.mail,
      ...GRAPH_SCOPES.calendar,
    ];

    const result = await pca.acquireTokenByRefreshToken({
      refreshToken: credential.refreshToken,
      scopes,
    });

    if (result) {
      // Save the refreshed credential
      const newCredential: Credential = {
        accessToken: result.accessToken,
        refreshToken: (result as any).refreshToken ?? credential.refreshToken,
        expiresAt: result.expiresOn?.toISOString() ?? new Date(Date.now() + 3600000).toISOString(),
        account: result.account?.username ?? credential.account,
        scopes: result.scopes?.length > 0 ? result.scopes : credential.scopes,
        clientId: credential.clientId,
        tenantId: credential.tenantId,
      };

      await setCredential("microsoft-graph", profile, newCredential);

      const expiresAt = new Date(newCredential.expiresAt);
      const expiresIn = Math.max(0, Math.floor((expiresAt.getTime() - Date.now()) / 1000));

      output({
        status: "valid",
        expiresIn,
        account: newCredential.account,
      });
      return;
    }

    output({ status: "needs-auth", error: "Refresh returned no result" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown error";
    output({ status: "needs-auth", error: `Refresh failed: ${message}` });
  }
}

main();
