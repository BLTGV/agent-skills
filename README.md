# API Skills

A Claude Code plugin providing skills for connecting to external APIs with guided access scripts.

## Features

- **Microsoft Graph API** - Access Microsoft 365 emails and calendar
- **Notion API** - *(Coming soon)*

## Prerequisites

- [Bun](https://bun.sh/) runtime
- Microsoft account (for Graph API)

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   bun install
   ```
3. Add to Claude Code:
   ```bash
   claude --plugin-dir /path/to/agent-skills
   ```

## Skills

### Microsoft Graph

Access Microsoft 365 email and calendar.

**Triggers:** "read my emails", "check my calendar", "search emails", "what meetings do I have"

#### Quick Start (Default Client)

Works immediately for most personal accounts:

```bash
bun run skills/microsoft-graph/scripts/auth.ts
```

#### Using Your Own App Registration (Recommended)

For work/school accounts or production use, register your own Azure AD app:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**:
   - **Name:** Your app name (e.g., "My Graph CLI")
   - **Supported account types:** Choose based on your needs:
     - "Accounts in this organizational directory only" (single tenant)
     - "Accounts in any organizational directory" (multi-tenant)
     - "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI:** Select "Public client/native (mobile & desktop)" and enter:
     ```
     https://login.microsoftonline.com/common/oauth2/nativeclient
     ```
4. Click **Register**
5. Go to **Authentication**:
   - Enable **Allow public client flows** (set to Yes)
   - Click **Save**
6. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**:
   - `Mail.Read` - Read user mail
   - `Calendars.Read` - Read user calendars
   - `User.Read` - Sign in and read user profile
   - Click **Add permissions**
7. Copy the **Application (client) ID** from the Overview page

Then authenticate with your app:

```bash
# Multi-tenant app
bun run skills/microsoft-graph/scripts/auth.ts --client-id YOUR_CLIENT_ID

# Single-tenant app (replace with your tenant ID)
bun run skills/microsoft-graph/scripts/auth.ts \
  --client-id YOUR_CLIENT_ID \
  --tenant-id YOUR_TENANT_ID
```

The client ID is saved with your credentials, so you only need to specify it once per profile.

#### Usage Examples

```bash
# List emails
bun run skills/microsoft-graph/scripts/emails.ts list

# Today's calendar
bun run skills/microsoft-graph/scripts/calendar.ts today

# Search emails
bun run skills/microsoft-graph/scripts/emails.ts search --query "from:boss@company.com"

# List mail folders
bun run skills/microsoft-graph/scripts/emails.ts folders
```

#### Multi-Account Support

```bash
# Add work account with custom app
bun run auth.ts --profile work --client-id YOUR_WORK_APP_ID

# Add personal account
bun run auth.ts --profile personal

# Use specific profile
bun run emails.ts list --profile work
bun run calendar.ts today --profile personal

# List all profiles
bun run auth.ts --list
```

## Credential Storage

Credentials are stored in `~/.config/api-skills/credentials.json`:

```json
{
  "microsoft-graph": {
    "default": {
      "accessToken": "...",
      "refreshToken": "...",
      "expiresAt": "...",
      "account": "user@example.com",
      "scopes": ["Mail.Read", "Calendars.Read"],
      "clientId": "your-app-id",
      "tenantId": "your-tenant-id"
    }
  }
}
```

- Tokens are automatically refreshed when expired
- Client ID and tenant ID are remembered per profile
- Different profiles can use different app registrations

## Troubleshooting

### "AADSTS50020: User account does not exist in tenant"

Your organization requires a single-tenant app. Register an app in your Azure AD tenant and use:
```bash
bun run auth.ts --client-id YOUR_APP_ID --tenant-id YOUR_TENANT_ID
```

### "AADSTS65001: User has not consented to use the application"

Your organization requires admin consent for apps. Ask your IT admin to:
1. Go to Azure AD > Enterprise applications > Your app
2. Grant admin consent for the required permissions

### "AADSTS7000218: Request body must contain client_assertion or client_secret"

Your app is configured as a confidential client. Go to Azure Portal > Your app > Authentication and enable "Allow public client flows".

## Adding New API Skills

This plugin is designed to be extensible. To add a new API integration:

1. Create a new skill directory: `skills/your-api/`
2. Add `SKILL.md` with triggers and usage documentation
3. Create TypeScript scripts in `scripts/`
4. Use the shared credential library for token storage

## Development

```bash
# Type check
bun run typecheck

# Run a script
bun run skills/microsoft-graph/scripts/auth.ts --help
```

## License

MIT
