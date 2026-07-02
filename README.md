# Microsoft Outlook MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Outlook email and calendar through the Microsoft Graph API.

[![Download Latest Release](https://img.shields.io/github/v/release/XenoXilus/outlook-mcp?label=Download&color=blue)](https://github.com/XenoXilus/outlook-mcp/releases/latest)
[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-ff5f5f?logo=ko-fi)](https://ko-fi.com/xenoxilus)

## Features

- **Email Operations**: Read, search, send, reply to emails and download attachments
- **SharePoint Integration**: Access SharePoint files via sharing links or direct file IDs. Download files shared to you via emails. 
- **Calendar Management**: View and manage calendar events and appointments
- **Office Document Processing**: Parse PDF, Word, PowerPoint, and Excel files with extracted text content
- **Multi-Account Support**: Connect multiple Microsoft accounts across tenants; search and calendar views fan out across all accounts
- **Zero-Config OAuth**: Connect accounts via `outlook_connect_account` without manual Azure app setup (BYO app optional)
- **Docker / Headless**: Run in a container with device-code OAuth — see [docs/docker.md](docs/docker.md)

## Quick Start

**Choose your installation method:**

| Method | Best For |
|--------|----------|
| [Docker](#docker-recommended) | Servers, CI, remote MCP gateways, headless environments |
| [DXT Extension](#installing-as-dxt-extension) | Claude Desktop users |
| [Node.js CLI](#using-with-cli-tools) | Local development with browser OAuth |

> **Getting started**: Install the server, then call `outlook_connect_account` to sign in. Azure app registration is optional — see [Azure Setup (Advanced)](#azure-setup-guide-advanced).

---

## Installation

### Docker (Recommended)

The recommended way to run outlook-mcp is in a Docker container. The image includes the server, uses **device-code OAuth** (no browser inside the container), and persists tokens on a mounted volume.

**Prerequisites:** [Docker](https://docs.docker.com/get-docker/) with Docker Compose

**1. Clone and build:**

```bash
git clone https://github.com/XOGENT/outlook-mcp.git
cd outlook-mcp
docker compose build
```

**2. Configure your MCP client:**

Add this to your MCP servers configuration (Claude Code `~/.claude.json`, project `.mcp.json`, etc.):

```json
{
  "outlook-mcp": {
    "command": "docker",
    "args": [
      "compose",
      "-f", "/absolute/path/to/outlook-mcp/docker-compose.yml",
      "run", "--rm", "-T", "outlook-mcp"
    ]
  }
}
```

Or use `docker run` directly (after `docker compose build` tags the image as `outlook-mcp:local`):

```json
{
  "outlook-mcp": {
    "command": "docker",
    "args": [
      "run", "-i", "--rm",
      "-v", "outlook-mcp-data:/data",
      "-e", "MCP_OUTLOOK_HEADLESS=true",
      "-e", "MCP_OUTLOOK_DATA_DIR=/data",
      "-e", "MCP_OUTLOOK_WORK_DIR=/data/downloads",
      "outlook-mcp:local"
    ]
  }
}
```

Replace `/absolute/path/to/outlook-mcp` with the path where you cloned the repo.

**3. Connect your account:**

1. Start your MCP client and call **`outlook_connect_account`**
2. The response includes a device login URL and code
3. Open https://microsoft.com/devicelogin on any device and enter the code
4. Repeat for additional accounts

**Manual test run:**

```bash
docker compose run --rm outlook-mcp
```

You should see `Outlook MCP server is ready and connected`.

**Data persistence:** Tokens and account registry are stored in the `outlook-mcp-data` Docker volume at `/data/accounts/`. Back up this volume to preserve connected accounts across container restarts.

See [docs/docker.md](docs/docker.md) for environment variables, security notes, and advanced deployment options.

---

### Installing as DXT Extension

For Claude Desktop users, DXT extensions provide the simplest installation experience.

**Option 1: Download Pre-built Extension**
1. Download `outlook-mcp.dxt` from the [Releases page](https://github.com/XenoXilus/outlook-mcp/releases)
2. In Claude Desktop, go to **Settings** → **Extensions**
3. Click **Install from file** and select the `.dxt` file
4. Click **Install from file** — Azure Client ID and Tenant ID are optional (Advanced)

**Connect your account:** After installing, ask your assistant to run `outlook_connect_account` to sign in via Microsoft OAuth.

**Option 2: Build from Source**
1. Clone and install dependencies:
   ```bash
   git clone https://github.com/XOGENT/outlook-mcp.git
   cd outlook-mcp
   npm install
   ```
2. Install the DXT CLI: `npm install -g @anthropic-ai/dxt`
3. Pack the extension:
   ```bash
   dxt pack . outlook-mcp.dxt
   ```
4. Install the generated `.dxt` file in Claude Desktop as above

---

### Using with CLI Tools (Node.js)

For local development or MCP clients that run Node.js directly on your machine (browser-based OAuth).

**1. Clone and Install:**
```bash
git clone https://github.com/XOGENT/outlook-mcp.git
cd outlook-mcp
npm install
```

**2. Configure your MCP client:**

Add the following to your MCP servers configuration (location varies by client):

```json
{
  "outlook-mcp": {
    "command": "node",
    "args": ["/absolute/path/to/outlook-mcp/server/index.js"],
    "env": {
      "MCP_OUTLOOK_DATA_DIR": "/absolute/path/to/outlook-mcp/.tokens",
      "MCP_OUTLOOK_WORK_DIR": "/absolute/path/to/outlook-mcp/.downloads"
    }
  }
}
```

`AZURE_CLIENT_ID` and `AZURE_TENANT_ID` are optional — omit them to use the default hosted OAuth app. Set them only if your organization requires a BYO Azure app (see [Azure Setup](#azure-setup-guide-advanced)).

**Common config file locations:**
- **Claude Code**: `~/.claude.json` or project-level `.mcp.json`
- **mcp CLI**: `~/.config/mcp/servers.json`

**3. Optional: Use environment variables**

Instead of specifying `env` in the config, you can export the variables in your shell:

```bash
export MCP_OUTLOOK_DATA_DIR="/absolute/path/to/outlook-mcp/.tokens"
export MCP_OUTLOOK_WORK_DIR="/absolute/path/to/outlook-mcp/.downloads"
# Optional BYO Azure app:
export AZURE_CLIENT_ID="your-azure-client-id"
export AZURE_TENANT_ID="your-azure-tenant-id"
```

---

## Multi-Account Setup

1. Call **`outlook_connect_account`** to sign in (browser on desktop, device code in Docker)
2. Repeat for each additional Microsoft account / tenant
3. Use **`outlook_list_accounts`** to see connected accounts
4. **Search & calendar list tools** query all accounts by default
5. **Send/write tools** require an `account` parameter when multiple accounts are connected
6. Use **`outlook_disconnect_account`** to remove an account

### Shared mailboxes

Pass the `mailbox` parameter (e.g. `billing@contoso.com`) to access shared/delegated mailboxes via `/users/{mailbox}` paths.

---

## Azure Setup Guide (Advanced)

Azure app registration is **optional**. Use this section if your organization requires a BYO app or blocks third-party consent.

To use the default hosted app, skip this section and call `outlook_connect_account` after install.

### For Business/Work Accounts (Recommended)

1. Go to the [Azure Portal](https://portal.azure.com/) and search for "App registrations".
2. Click **New registration**.
   - Name: `Outlook MCP` (or similar)
   - Supported account types: **Accounts in this organizational directory only** (Single tenant)
   - Redirect URI: Select **Web** and enter `http://localhost/callback`
3. Click **Register**.
4. Go to **Authentication** in the sidebar.
   - Under "Advanced settings", set **Allow public client flows** to **Yes**.
   - Click **Save**.
5. On the Overview page, copy:
   - **Application (client) ID** → This is your `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → This is your `AZURE_TENANT_ID`
6. Go to **API permissions** in the sidebar.
   - Click **Add a permission** -> **Microsoft Graph** -> **Delegated permissions**.
   - Add these permissions:
     - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
     - `Calendars.Read`, `Calendars.ReadWrite`
     - `User.Read`, `MailboxSettings.Read`
     - `Files.Read.All`, `Files.ReadWrite.All`
     - `Sites.Read.All`, `Sites.ReadWrite.All`
     - `offline_access`
   - Click **Add permissions**.
   - (Optional) If you are an admin, click **Grant admin consent** to suppress consent prompts for users.

**Note:** No client secret is required (PKCE auth flow).

### For Personal Accounts (outlook.com, hotmail.com)

Personal Microsoft accounts can also register apps in Azure:

1. Sign in to the [Azure Portal](https://portal.azure.com/) with your personal Microsoft account (outlook.com, hotmail.com, etc.).
2. If prompted to create a directory, follow the steps to create a free Azure directory.
3. Follow the same steps as above for Business accounts.
4. When configuring, use **Accounts in any organizational directory and personal Microsoft accounts** for supported account types.

---

## Configuration Reference

### Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `MCP_OUTLOOK_DATA_DIR` | No | `.tokens` (local) / `/data` (Docker) | Token storage and account registry directory |
| `MCP_OUTLOOK_WORK_DIR` | No | `.downloads` (local) / `/data/downloads` (Docker) | Directory for large attachment downloads |
| `MCP_OUTLOOK_HEADLESS` | No | `true` in Docker image | Use device-code OAuth instead of browser flow |
| `AZURE_CLIENT_ID` | No | bundled default | BYO Azure AD application client ID |
| `AZURE_TENANT_ID` | No | `organizations` | BYO Azure AD tenant ID |

### Large File Handling

When downloading large attachments or SharePoint files, the server automatically detects when the response would exceed the MCP 1MB limit and saves the content to local files instead.

- In Docker, files are saved to `MCP_OUTLOOK_WORK_DIR` (`/data/downloads` by default)
- Locally, files are saved to `MCP_OUTLOOK_WORK_DIR` or the system temp directory
- Files are automatically named with timestamps to avoid conflicts
- Old files are periodically cleaned up to manage disk space

---

## Example Prompts

Once installed, you can ask the AI assistant things like:

**Email Management**
- "Show me my unread emails from this week"
- "Find all emails from John about the project proposal"
- "Send a reply to the last email from Sarah thanking her for the update"
- "Draft an email to the team summarizing today's meeting"

**Calendar**
- "What meetings do I have tomorrow?"
- "Schedule a 30-minute call with Alex next Tuesday afternoon"
- "Show me my availability for the rest of the week"

**Attachments & SharePoint**
- "Download and summarize the PDF attachment from the latest email from Finance"
- "Get the contents of this SharePoint link: [paste link]"
- "What files were attached to emails from Legal this month?"

**Office Document Processing**

The server automatically parses:
- **PDF files**: Extracts text content
- **Word documents** (.docx): Extracts text content
- **PowerPoint** (.pptx): Extracts slide text
- **Excel** (.xlsx): Parses data into structured format

---

## Authentication

The server uses OAuth 2.0 for secure authentication:

| Environment | OAuth flow | How to sign in |
|-------------|------------|----------------|
| **Docker / headless** | Device code | Call `outlook_connect_account`, then visit https://microsoft.com/devicelogin with the code from the response |
| **Local Node.js** | PKCE (browser) | Call `outlook_connect_account` — a browser window opens for Microsoft sign-in |

**Token storage:**
1. Tokens are encrypted and stored under `MCP_OUTLOOK_DATA_DIR` (Docker volume at `/data`, local default `.tokens`)
2. Per-account storage supports multiple connected Microsoft accounts
3. Automatic token refresh for long-term usage
4. OS keychain is used when available locally; Docker always uses encrypted file storage

### Required Permissions

The app requests these Microsoft Graph permissions:

- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` - Email access
- `Calendars.Read`, `Calendars.ReadWrite` - Calendar access  
- `User.Read`, `MailboxSettings.Read` - User profile
- `Files.Read.All`, `Files.ReadWrite.All` - OneDrive/SharePoint files
- `Sites.Read.All`, `Sites.ReadWrite.All` - SharePoint sites
- `offline_access` - Refresh tokens

---

## Troubleshooting

### Docker Issues
- **Problem**: `EACCES: permission denied, mkdir '/app/.tokens'`
- **Solution**: Rebuild the image (`docker compose build`) and ensure the data volume is mounted. The server must write to `/data`, not `/app`.
- **Problem**: Container exits immediately
- **Solution**: MCP clients need stdio — use `docker compose run --rm -T` or the `docker run -i` config above. Do not use `docker compose up` without an attached MCP client.

### Large File Issues
- **Problem**: "Result exceeds maximum length" error
- **Solution**: Ensure `MCP_OUTLOOK_WORK_DIR` is set and writable
- **Alternative**: Files automatically save to system temp if work dir not configured

### Authentication Issues
- **Problem**: Authentication failures
- **Solution**: Verify Azure AD app permissions if using a BYO app; otherwise call `outlook_connect_account` again
- **Docker**: Complete device-code sign-in at https://microsoft.com/devicelogin before retrying tools
- **Reset**: Remove the Docker volume (`docker volume rm outlook-mcp-data`) or delete `MCP_OUTLOOK_DATA_DIR` locally, then reconnect

### SharePoint Access Issues
- **Problem**: Cannot access SharePoint files
- **Solution**: Ensure sharing links are valid and user has access permissions
- **Alternative**: Use direct file ID access if available

---

## Development

### Project Structure
```
outlook-mcp/
├── Dockerfile              # Container image for headless deployment
├── docker-compose.yml      # Local Docker setup with persistent volume
├── server/
│   ├── index.js              # Main MCP server
│   ├── auth/                 # Authentication management
│   ├── graph/                # Microsoft Graph API client
│   ├── schemas/              # MCP tool schemas
│   ├── tools/                # MCP tool implementations
│   │   ├── attachments/      # Attachment tools
│   │   ├── calendar/         # Calendar tools
│   │   ├── email/            # Email tools
│   │   ├── folders/          # Folder management
│   │   └── sharepoint/       # SharePoint tools
│   └── utils/                # Utility modules
└── package.json
```

### Running Tests
```bash
npm test                    # Run all tests
npm run test:watch          # Watch mode
npm run test:benchmark      # Performance benchmarks
```

### Debugging
```bash
npm run test:graph          # Test Graph API connection
```

---

## Support

If this tool saved you time, consider supporting the development!

[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-ff5f5f?logo=ko-fi)](https://ko-fi.com/xenoxilus)

---

## License

MIT License

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with tests
4. Submit a pull request
