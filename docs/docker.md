# Docker Headless Deployment

Run outlook-mcp in a container without a browser. OAuth uses the **device code flow** — the MCP tool response includes a URL and code for the user to complete sign-in on any device.

## Quick Start

```bash
docker compose build
docker compose run --rm outlook-mcp
```

For MCP clients, configure stdio transport:

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

## First-Time Setup

1. Start the MCP server (container or local with `MCP_OUTLOOK_HEADLESS=true`)
2. Call `outlook_connect_account` — the response includes a device login URL and code
3. Complete sign-in at https://microsoft.com/devicelogin
4. Repeat `outlook_connect_account` for additional accounts

## Environment Variables

| Variable | Default (Docker) | Purpose |
|----------|------------------|---------|
| `MCP_OUTLOOK_DATA_DIR` | `/data` | Token and account registry persistence |
| `MCP_OUTLOOK_HEADLESS` | `true` | Force device code OAuth flow |
| `MCP_OUTLOOK_WORK_DIR` | `/data/downloads` | Attachment download directory |
| `AZURE_CLIENT_ID` | bundled default | Optional BYO app client ID |
| `AZURE_TENANT_ID` | unset | Optional BYO tenant ID |

## Data Persistence

Tokens are stored under `/data/accounts/` in the mounted volume. Back up this volume to preserve connected accounts across container restarts.

## Security

- Mount the data volume with restricted filesystem permissions
- Do not bake secrets into the image
- The container runs as the `node` user
