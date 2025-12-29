---
name: Microsoft Graph API
description: This skill should be used when the user asks to "read my emails", "check my calendar", "get calendar events", "search emails", "list mail folders", "show unread messages", "what meetings do I have", "fetch emails from Microsoft", "access Outlook", or mentions Microsoft Graph, Office 365 email, or Outlook calendar integration.
version: 0.1.0
---

# Microsoft Graph API Integration

Access Microsoft 365 emails and calendar through TypeScript scripts executed via Bun. The scripts handle OAuth authentication, token management, and API requests.

## Overview

This skill provides guided access to Microsoft Graph API for:
- **Email**: List, read, and search emails across folders
- **Calendar**: View and search calendar events
- **Authentication**: Device code OAuth flow with multi-profile support

All scripts are located at `${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/`.

## Authentication

Before accessing emails or calendar, authenticate using the device code flow:

```bash
bun run ${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/auth.ts
```

The script displays a URL and code. Open the URL in a browser, enter the code, and sign in with a Microsoft account. Credentials are stored in `~/.config/api-skills/credentials.json`.

### Multi-Profile Support

Store multiple accounts using profiles:

```bash
# Authenticate with work account
bun run auth.ts --profile work

# Authenticate with personal account
bun run auth.ts --profile personal

# List all profiles
bun run auth.ts --list

# Delete a profile
bun run auth.ts --delete --profile old-account
```

### Credential Storage

Credentials are stored at `~/.config/api-skills/credentials.json`:

```json
{
  "microsoft-graph": {
    "default": {
      "accessToken": "...",
      "refreshToken": "...",
      "expiresAt": "2024-01-15T10:30:00.000Z",
      "account": "user@example.com",
      "scopes": ["Mail.Read", "Calendars.Read", "User.Read"]
    }
  }
}
```

Tokens are automatically refreshed when expired.

### Authentication Recovery Flow

When running email or calendar commands, if authentication fails, use this recovery flow:

1. **Check auth status:**
   ```bash
   bun run ${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/check-auth.ts --json
   ```

   If status is `"valid"`, proceed with commands normally.

2. **If status is `"needs-auth"`, start device code flow.**

   **IMPORTANT:** Do NOT run this in the background. The script exits immediately with `--no-wait` and outputs JSON that MUST be shown to the user.

   ```bash
   bun run ${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/auth.ts --json --no-wait
   ```

   Output:
   ```json
   {"userCode":"ABC123","verificationUri":"https://microsoft.com/devicelogin","expiresIn":900,"message":"..."}
   ```

3. **IMMEDIATELY display to user** the `userCode` and `verificationUri` from the JSON output:
   ```
   To access your email, please authenticate:
   1. Go to: https://microsoft.com/devicelogin
   2. Enter code: ABC123

   Let me know when you've completed authentication.
   ```

4. **When user confirms, verify:**
   ```bash
   bun run ${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/check-auth-complete.ts --json
   ```

5. **Retry the original command** if authentication is complete.

### Token Lifecycle

| Token Type | Lifetime | Handling |
|------------|----------|----------|
| Access Token | ~1 hour | Automatically refreshed using refresh token |
| Refresh Token | ~90 days | When expired, requires device code flow |

The `check-auth.ts` script attempts silent refresh automatically. Users only need to re-authenticate when the refresh token itself has expired (roughly every 90 days).

## Email Access

Use `emails.ts` to interact with emails:

### List Emails

```bash
# List inbox (default)
bun run ${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/emails.ts list

# List from specific folder
bun run emails.ts list --folder "Sent Items"
bun run emails.ts list --folder drafts --top 5

# Use different profile
bun run emails.ts list --profile work
```

### Read Specific Email

```bash
bun run emails.ts read --id AAMkAG...
```

Get the ID from the `list` command output.

### Search Emails

```bash
# Search by sender
bun run emails.ts search --query "from:boss@company.com"

# Search by subject
bun run emails.ts search --query "subject:quarterly report"

# Combined search
bun run emails.ts search --query "from:hr@company.com subject:benefits"

# Search with attachments
bun run emails.ts search --query "hasAttachments:true"
```

### List Mail Folders

```bash
bun run emails.ts folders
```

Shows folder names, IDs, and message counts.

### JSON Output

Add `--format json` for machine-readable output:

```bash
bun run emails.ts list --format json
```

## Calendar Access

Use `calendar.ts` to view calendar events:

### List Upcoming Events

```bash
# Default: next 30 days
bun run ${CLAUDE_PLUGIN_ROOT}/skills/microsoft-graph/scripts/calendar.ts list

# Today's events
bun run calendar.ts today

# This week's events
bun run calendar.ts week

# Custom date range
bun run calendar.ts list --start tomorrow --end +7d
bun run calendar.ts list --start 2024-02-01 --end 2024-02-28
```

### Date Formats

- **Relative**: `today`, `tomorrow`, `+7d`, `+1m`, `+1y`
- **Absolute**: ISO format `2024-01-15`

### View Specific Event

```bash
bun run calendar.ts view --id AAMkAG...
```

Shows full details including attendees and description.

### Search Events

```bash
bun run calendar.ts search --query "team standup"
bun run calendar.ts search --query "1:1"
```

## Common Workflows

### Check for Important Emails

```bash
# Unread from specific sender
bun run emails.ts search --query "from:ceo@company.com isRead:false"

# Recent urgent emails
bun run emails.ts search --query "subject:urgent" --top 5
```

### Review Today's Schedule

```bash
bun run calendar.ts today
```

### Find Meeting Details

```bash
# Search for meeting
bun run calendar.ts search --query "project kickoff"

# Get full details with attendees
bun run calendar.ts view --id <event-id>
```

### Check Multiple Accounts

```bash
# Personal inbox
bun run emails.ts list --profile personal

# Work calendar
bun run calendar.ts today --profile work
```

## Error Handling

### Authentication Errors

If token is expired or invalid:
```
Error: No valid token for profile "default". Run auth first:
bun run auth.ts --profile default
```

Re-run authentication to refresh credentials.

### API Errors

Common Graph API errors:
- **401 Unauthorized**: Token expired, re-authenticate
- **403 Forbidden**: Missing required scopes, re-auth with correct scopes
- **404 Not Found**: Invalid folder name or message ID

### Scope Requirements

Different operations require different scopes:
- **Email**: `Mail.Read` or `Mail.ReadBasic`
- **Calendar**: `Calendars.Read`
- **User info**: `User.Read`

To request specific scopes:
```bash
bun run auth.ts --scopes Mail.Read,Calendars.Read
```

## Script Reference

All scripts support `--help` for detailed usage:

| Script | Purpose |
|--------|---------|
| `auth.ts` | OAuth authentication and credential management |
| `check-auth.ts` | Check auth status and attempt silent refresh |
| `check-auth-complete.ts` | Verify authentication completed |
| `emails.ts` | Email list, read, search, and folder operations |
| `calendar.ts` | Calendar view and search operations |

## Additional Resources

For detailed API reference and advanced patterns, see:
- **`references/graph-api.md`** - Microsoft Graph API endpoints and parameters
