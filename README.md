<div align="center">

# <span style="color:#cad8d9">Google Workspace MCP Server</span> <img src="https://github.com/user-attachments/assets/b89524e4-6e6e-49e6-ba77-00d6df0c6e5c" width="80" align="right" />

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![PyPI](https://img.shields.io/pypi/v/workspace-mcp.svg)](https://pypi.org/project/workspace-mcp/)
[![PyPI Downloads](https://static.pepy.tech/badge/workspace-mcp/month)](https://pepy.tech/projects/workspace-mcp)
[![Website](https://img.shields.io/badge/Website-workspacemcp.com-green.svg)](https://workspacemcp.com)

*Full natural language control over Google Calendar, Drive, Gmail, Docs, Sheets, Slides, Forms, Tasks, and Chat through all MCP clients, AI assistants and developer tools.*

**The most feature-complete Google Workspace MCP server**, now with Remote OAuth2.1 multi-user support and 1-click Claude installation.


###### Support for all free Google accounts (Gmail, Docs, Drive etc) & Google Workspace plans (Starter, Standard, Plus, Enterprise, Non Profit) with expanded app options like Chat & Spaces. <br/> Interested in a private cloud instance? [That can be arranged.](https://workspacemcp.com/workspace-mcp-cloud)


</div>

<div align="center">
<a href="https://glama.ai/mcp/servers/@taylorwilsdon/google_workspace_mcp">
  <img width="195" src="https://glama.ai/mcp/servers/@taylorwilsdon/google_workspace_mcp/badge" alt="Google Workspace Server MCP server" align="center"/>
</a>
<a href="https://www.pulsemcp.com/servers/taylorwilsdon-google-workspace">
<img width="456" src="https://github.com/user-attachments/assets/0794ef1a-dc1c-447d-9661-9c704d7acc9d" align="center"/>
</a>
</div>

---


**See it in action:**
<div align="center">
  <video width="400" src="https://github.com/user-attachments/assets/a342ebb4-1319-4060-a974-39d202329710"></video>
</div>

---

### A quick plug for AI-Enhanced Docs
<details>
<summary>◆ <b>But why?</b></summary>

**This README was written with AI assistance, and here's why that matters**
>
> As a solo dev building open source tools, comprehensive documentation often wouldn't happen without AI help. Using agentic dev tools like **Roo** & **Claude Code** that understand the entire codebase, AI doesn't just regurgitate generic content - it extracts real implementation details and creates accurate, specific documentation.
>
> In this case, Sonnet 4 took a pass & a human (me) verified them 8/16/25.
</details>

## <span style="color:#adbcbc">Overview</span>

A production-ready MCP server that integrates all major Google Workspace services with AI assistants. It supports both single-user operation and multi-user authentication via OAuth 2.1, making it a powerful backend for custom applications. Built with FastMCP for optimal performance, featuring advanced authentication handling, service caching, and streamlined development patterns.

**Simplified Setup**: Now uses Google Desktop OAuth clients - no redirect URIs or port configuration needed!

## <span style="color:#adbcbc">Features</span>

<table align="center" style="width: 100%; max-width: 100%;">
<tr>
<td width="50%" valign="top">

**<span style="color:#72898f">@</span> Gmail** • **<span style="color:#72898f">≡</span> Drive** • **<span style="color:#72898f">⧖</span> Calendar** **<span style="color:#72898f">≡</span> Docs**
- Complete Gmail management, end to end coverage
- Full calendar management with advanced features
- File operations with Office format support
- Document creation, editing & comments
- Deep, exhaustive support for fine grained editing

---

**<span style="color:#72898f">≡</span> Forms** • **<span style="color:#72898f">@</span> Chat** • **<span style="color:#72898f">≡</span> Sheets** • **<span style="color:#72898f">≡</span> Slides**
- Form creation, publish settings & response management
- Space management & messaging capabilities
- Spreadsheet operations with flexible cell management
- Presentation creation, updates & content manipulation

</td>
<td width="50%" valign="top">

**<span style="color:#72898f">⊠</span> Authentication & Security**
- Advanced OAuth 2.0 & OAuth 2.1 support
- Automatic token refresh & session management
- Transport-aware callback handling
- Multi-user bearer token authentication
- Innovative CORS proxy architecture

---

**<span style="color:#72898f">✓</span> Tasks** • **<span style="color:#72898f">◆</span> Custom Search** • **<span style="color:#72898f">↻</span> Transport Support**
- Full support for all MCP Transports
- OpenAPI compatibility via `mcpo`
- Task & task list management with hierarchy
- Programmable Search Engine (PSE) integration

</td>
</tr>
</table>

---

## ▶ Quick Start

<details>
<summary>≡ <b>Quick Reference Card</b> <sub><sup>← Essential commands & configs at a glance</sup></sub></summary>

<table>
<tr><td width="33%" valign="top">

**⊠ Credentials**
```bash
export GOOGLE_OAUTH_CLIENT_ID="..."
export GOOGLE_OAUTH_CLIENT_SECRET="..."
```
[Full setup →](#-credential-configuration)

</td><td width="33%" valign="top">

**▶ Launch Commands**
```bash
uvx workspace-mcp --tool-tier core
uv run main.py --tools gmail drive
```
[More options →](#start-the-server)

</td><td width="34%" valign="top">

**★ Tool Tiers**
- ● `core` - Essential tools
- ◐ `extended` - Core + extras
- ○ `complete` - Everything
[Details →](#tool-tiers)

</td></tr>
</table>

</details>

### 1. One-Click Claude Desktop Install (Recommended)

1. **Download:** Grab the latest `google_workspace_mcp.dxt` from the “Releases” page
2. **Install:** Double-click the file – Claude Desktop opens and prompts you to **Install**
3. **Configure:** In Claude Desktop → **Settings → Extensions → Google Workspace MCP**, paste your Google OAuth credentials
4. **Use it:** Start a new Claude chat and call any Google Workspace tool

>
**Why DXT?**
> Desktop Extensions (`.dxt`) bundle the server, dependencies, and manifest so users go from download → working MCP in **one click** – no terminal, no JSON editing, no version conflicts.

#### Required Configuration
<details>
<summary>◆ <b>Environment Variables</b> <sub><sup>← Click to configure in Claude Desktop</sup></sub></summary>

<table>
<tr><td width="50%" valign="top">

**Required**
| Variable | Purpose |
|----------|---------|
| `GOOGLE_OAUTH_CLIENT_ID` | OAuth client ID from Google Cloud |
| `GOOGLE_OAUTH_CLIENT_SECRET` | OAuth client secret |
| `OAUTHLIB_INSECURE_TRANSPORT=1` | Development only (allows `http://` redirect) |

</td><td width="50%" valign="top">

**Optional**
| Variable | Purpose |
|----------|---------|
| `USER_GOOGLE_EMAIL` | Default email for single-user auth |
| `GOOGLE_PSE_API_KEY` | API key for Custom Search |
| `GOOGLE_PSE_ENGINE_ID` | Search Engine ID for Custom Search |
| `MCP_ENABLE_OAUTH21` | Set to `true` for OAuth 2.1 support |

</td></tr>
</table>

Claude Desktop stores these securely in the OS keychain; set them once in the extension pane.
</details>

---

<div align="center">
  <video width="832" src="https://github.com/user-attachments/assets/83cca4b3-5e94-448b-acb3-6e3a27341d3a"></video>
</div>

---

### Prerequisites

- **Python 3.10+**
- **[uvx](https://github.com/astral-sh/uv)** (for instant installation) or [uv](https://github.com/astral-sh/uv) (for development)
- **Google Cloud Project** with OAuth 2.0 credentials

### Configuration

<details open>
<summary>◆ <b>Google Cloud Setup</b> <sub><sup>← OAuth 2.0 credentials & API enablement</sup></sub></summary>

<table>
<tr>
<td width="33%" align="center">

**1. Create Project**
```text
console.cloud.google.com

→ Create new project
→ Note project name
```
<sub>[Open Console →](https://console.cloud.google.com/)</sub>

</td>
<td width="33%" align="center">

**2. OAuth Credentials**
```text
APIs & Services → Credentials
→ Create Credentials
→ OAuth Client ID
→ Desktop Application
```
<sub>Download & save credentials</sub>

</td>
<td width="34%" align="center">

**3. Enable APIs**
```text
APIs & Services → Library

Search & enable:
Calendar, Drive, Gmail,
Docs, Sheets, Slides,
Forms, Tasks, Chat, Search
```
<sub>See quick links below</sub>

</td>
</tr>
<tr>
<td colspan="3">

<details>
<summary>≡ <b>OAuth Credential Setup Guide</b> <sub><sup>← Step-by-step instructions</sup></sub></summary>

**Complete Setup Process:**

1. **Create OAuth 2.0 Credentials** - Visit [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project (or use existing)
   - Navigate to **APIs & Services → Credentials**
   - Click **Create Credentials → OAuth Client ID**
   - Choose **Desktop Application** as the application type (no redirect URIs needed!)
   - Download credentials and note the Client ID and Client Secret

2. **Enable Required APIs** - In **APIs & Services → Library**
   - Search for and enable each required API
   - Or use the quick links below for one-click enabling

3. **Configure Environment** - Set your credentials:
   ```bash
   export GOOGLE_OAUTH_CLIENT_ID="your-client-id"
   export GOOGLE_OAUTH_CLIENT_SECRET="your-secret"
   ```

≡ [Full Documentation →](https://developers.google.com/workspace/guides/auth-overview)

</details>

</td>
</tr>
</table>

<details>
  <summary>⊥ <b>Quick API Enable Links</b> <sub><sup>← One-click enable each Google API</sup></sub></summary>
  You can enable each one by clicking the links below (make sure you're logged into the Google Cloud Console and have the correct project selected):

* [Enable Google Calendar API](https://console.cloud.google.com/flows/enableapi?apiid=calendar-json.googleapis.com)
* [Enable Google Drive API](https://console.cloud.google.com/flows/enableapi?apiid=drive.googleapis.com)
* [Enable Gmail API](https://console.cloud.google.com/flows/enableapi?apiid=gmail.googleapis.com)
* [Enable Google Docs API](https://console.cloud.google.com/flows/enableapi?apiid=docs.googleapis.com)
* [Enable Google Sheets API](https://console.cloud.google.com/flows/enableapi?apiid=sheets.googleapis.com)
* [Enable Google Slides API](https://console.cloud.google.com/flows/enableapi?apiid=slides.googleapis.com)
* [Enable Google Forms API](https://console.cloud.google.com/flows/enableapi?apiid=forms.googleapis.com)
* [Enable Google Tasks API](https://console.cloud.google.com/flows/enableapi?apiid=tasks.googleapis.com)
* [Enable Google Chat API](https://console.cloud.google.com/flows/enableapi?apiid=chat.googleapis.com)
* [Enable Google Custom Search API](https://console.cloud.google.com/flows/enableapi?apiid=customsearch.googleapis.com)

</details>

</details>

1.1. **Credentials**: See [Credential Configuration](#credential-configuration) for detailed setup options

2. **Environment Configuration**:

<details open>
<summary>◆ <b>Environment Variables</b> <sub><sup>← Configure your runtime environment</sup></sub></summary>

<table>
<tr>
<td width="33%" align="center">

**◆ Development Mode**
```bash
export OAUTHLIB_INSECURE_TRANSPORT=1
```
<sub>Allows HTTP redirect URIs</sub>

</td>
<td width="33%" align="center">

**@ Default User**
```bash
export USER_GOOGLE_EMAIL=\
  your.email@gmail.com
```
<sub>Single-user authentication</sub>

</td>
<td width="34%" align="center">

**◆ Custom Search**
```bash
export GOOGLE_PSE_API_KEY=xxx
export GOOGLE_PSE_ENGINE_ID=yyy
```
<sub>Optional: Search API setup</sub>

</td>
</tr>
</table>

</details>

3. **Server Configuration**:

<details open>
<summary>◆ <b>Server Settings</b> <sub><sup>← Customize ports, URIs & proxies</sup></sub></summary>

<table>
<tr>
<td width="33%" align="center">

**◆ Base Configuration**
```bash
export WORKSPACE_MCP_BASE_URI=
  http://localhost
export WORKSPACE_MCP_PORT=8000
```
<sub>Server URL & port settings</sub>

</td>
<td width="33%" align="center">

**↻ Proxy Support**
```bash
export MCP_ENABLE_OAUTH21=
  true
```
<sub>Leverage multi-user OAuth2.1 clients</sub>

</td>
<td width="34%" align="center">

**@ Default Email**
```bash
export USER_GOOGLE_EMAIL=\
  your.email@gmail.com
```
<sub>Skip email in auth flows in single user mode</sub>

</td>
</tr>
</table>

<details>
<summary>≡ <b>Configuration Details</b> <sub><sup>← Learn more about each setting</sup></sub></summary>

| Variable | Description | Default |
|----------|-------------|---------|
| `WORKSPACE_MCP_BASE_URI` | Base server URI (no port) | `http://localhost` |
| `WORKSPACE_MCP_PORT` | Server listening port | `8000` |
| `WORKSPACE_EXTERNAL_URL` | External URL for reverse proxy setups | None |
| `GOOGLE_OAUTH_REDIRECT_URI` | Override OAuth callback URL | Auto-constructed |
| `USER_GOOGLE_EMAIL` | Default auth email | None |

</details>

</details>

### Google Custom Search Setup

<details>
<summary>◆ <b>Custom Search Configuration</b> <sub><sup>← Enable web search capabilities</sup></sub></summary>

<table>
<tr>
<td width="33%" align="center">

**1. Create Search Engine**
```text
programmablesearchengine.google.com
/controlpanel/create

→ Configure sites or entire web
→ Note your Engine ID (cx)
```
<sub>[Open Control Panel →](https://programmablesearchengine.google.com/controlpanel/create)</sub>

</td>
<td width="33%" align="center">

**2. Get API Key**
```text
developers.google.com
/custom-search/v1/overview

→ Create/select project
→ Enable Custom Search API
→ Create credentials (API Key)
```
<sub>[Get API Key →](https://developers.google.com/custom-search/v1/overview)</sub>

</td>
<td width="34%" align="center">

**3. Set Variables**
```bash
export GOOGLE_PSE_API_KEY=\
  "your-api-key"
export GOOGLE_PSE_ENGINE_ID=\
  "your-engine-id"
```
<sub>Configure in environment</sub>

</td>
</tr>
<tr>
<td colspan="3">

<details>
<summary>≡ <b>Quick Setup Guide</b> <sub><sup>← Step-by-step instructions</sup></sub></summary>

**Complete Setup Process:**

1. **Create Search Engine** - Visit the [Control Panel](https://programmablesearchengine.google.com/controlpanel/create)
   - Choose "Search the entire web" or specify sites
   - Copy the Search Engine ID (looks like: `017643444788157684527:6ivsjbpxpqw`)

2. **Enable API & Get Key** - Visit [Google Developers Console](https://console.cloud.google.com/)
   - Enable "Custom Search API" in your project
   - Create credentials → API Key
   - Restrict key to Custom Search API (recommended)

3. **Configure Environment** - Add to your shell or `.env`:
   ```bash
   export GOOGLE_PSE_API_KEY="AIzaSy..."
   export GOOGLE_PSE_ENGINE_ID="01764344478..."
   ```

≡ [Full Documentation →](https://developers.google.com/custom-search/v1/overview)

</details>

</td>
</tr>
</table>

</details>

### Start the Server

<details open>
<summary>▶ <b>Launch Commands</b> <sub><sup>← Choose your startup mode</sup></sub></summary>

<table>
<tr>
<td width="33%" align="center">

**▶ Quick Start**
```bash
uv run main.py
```
<sub>Default stdio mode</sub>

</td>
<td width="33%" align="center">

**◆ HTTP Mode**
```bash
uv run main.py \
  --transport streamable-http
```
<sub>Web interfaces & debugging</sub>

</td>
<td width="34%" align="center">

**@ Single User**
```bash
uv run main.py \
  --single-user
```
<sub>Simplified authentication</sub>

</td>
</tr>
<tr>
<td colspan="3">

<details>
<summary>◆ <b>Advanced Options</b> <sub><sup>← Tool selection, tiers & Docker</sup></sub></summary>

**▶ Selective Tool Loading**
```bash
# Load specific services only
uv run main.py --tools gmail drive calendar
uv run main.py --tools sheets docs

# Combine with other flags
uv run main.py --single-user --tools gmail
```

**★ Tool Tiers**
```bash
uv run main.py --tool-tier core      # ● Essential tools only
uv run main.py --tool-tier extended  # ◐ Core + additional
uv run main.py --tool-tier complete  # ○ All available tools
```

**◆ Docker Deployment**
```bash
docker build -t workspace-mcp .
docker run -p 8000:8000 -v $(pwd):/app \
  workspace-mcp --transport streamable-http

# With tool selection via environment variables
docker run -e TOOL_TIER=core workspace-mcp
docker run -e TOOLS="gmail drive calendar" workspace-mcp
```

**Available Services**: `gmail` • `drive` • `calendar` • `docs` • `sheets` • `forms` • `tasks` • `chat` • `search`

</details>

</td>
</tr>
</table>

</details>

### Tool Tiers

The server organizes tools into **three progressive tiers** for simplified deployment. Choose a tier that matches your usage needs and API quota requirements.

<table>
<tr>
<td width="65%" valign="top">

#### <span style="color:#72898f">Available Tiers</span>

**<span style="color:#2d5b69">●</span> Core** (`--tool-tier core`)
Essential tools for everyday tasks. Perfect for light usage with minimal API quotas. Includes search, read, create, and basic modify operations across all services.

**<span style="color:#72898f">●</span> Extended** (`--tool-tier extended`)
Core functionality plus management tools. Adds labels, folders, batch operations, and advanced search. Ideal for regular usage with moderate API needs.

**<span style="color:#adbcbc">●</span> Complete** (`--tool-tier complete`)
Full API access including comments, headers/footers, publishing settings, and administrative functions. For power users needing maximum functionality.

</td>
<td width="35%" valign="top">

#### <span style="color:#72898f">Important Notes</span>

<span style="color:#72898f">▶</span> **Start with `core`** and upgrade as needed
<span style="color:#72898f">▶</span> **Tiers are cumulative** – each includes all previous
<span style="color:#72898f">▶</span> **Mix and match** with `--tools` for specific services
<span style="color:#72898f">▶</span> **Configuration** in `core/tool_tiers.yaml`
<span style="color:#72898f">▶</span> **Authentication** included in all tiers

</td>
</tr>
</table>

#### <span style="color:#72898f">Usage Examples</span>

```bash
# Basic tier selection
uv run main.py --tool-tier core                            # Start with essential tools only
uv run main.py --tool-tier extended                        # Expand to include management features
uv run main.py --tool-tier complete                        # Enable all available functionality

# Selective service loading with tiers
uv run main.py --tools gmail drive --tool-tier core        # Core tools for specific services
uv run main.py --tools gmail --tool-tier extended          # Extended Gmail functionality only
uv run main.py --tools docs sheets --tool-tier complete    # Full access to Docs and Sheets
```

## 📋 Credential Configuration

<details open>
<summary>🔑 <b>OAuth Credentials Setup</b> <sub><sup>← Essential for all installations</sup></sub></summary>

<table>
<tr>
<td width="33%" align="center">

**🚀 Environment Variables**
```bash
export GOOGLE_OAUTH_CLIENT_ID=\
  "your-client-id"
export GOOGLE_OAUTH_CLIENT_SECRET=\
  "your-secret"
```
<sub>Best for production</sub>

</td>
<td width="33%" align="center">

**📁 File-based**
```bash
# Download & place in project root
client_secret.json

# Or specify custom path
export GOOGLE_CLIENT_SECRET_PATH=\
  /path/to/secret.json
```
<sub>Traditional method</sub>

</td>
<td width="34%" align="center">

**⚡ .env File**
```bash
cp .env.oauth21 .env
# Edit .env with credentials
```
<sub>Best for development</sub>

</td>
</tr>
<tr>
<td colspan="3">

<details>
<summary>📖 <b>Credential Loading Details</b> <sub><sup>← Understanding priority & best practices</sup></sub></summary>

**Loading Priority**
1. Environment variables (`export VAR=value`)
2. `.env` file in project root
3. `client_secret.json` via `GOOGLE_CLIENT_SECRET_PATH`
4. Default `client_secret.json` in project root

**Why Environment Variables?**
- ✅ **Docker/K8s ready** - Native container support
- ✅ **Cloud platforms** - Heroku, Railway, Vercel
- ✅ **CI/CD pipelines** - GitHub Actions, Jenkins
- ✅ **No secrets in git** - Keep credentials secure
- ✅ **Easy rotation** - Update without code changes

</details>

</td>
</tr>
</table>

</details>

---

## 🧰 Available Tools

> **Note**: All tools support automatic authentication via `@require_google_service()` decorators with 30-minute service caching.

<table width="100%">
<tr>
<td width="50%" valign="top">

### 📅 **Google Calendar** <sub>[`calendar_tools.py`](gcalendar/calendar_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `list_calendars` | **Core** | List accessible calendars |
| `get_events` | **Core** | Retrieve events with time range filtering |
| `create_event` | **Core** | Create events with attachments & reminders |
| `modify_event` | **Core** | Update existing events |
| `delete_event` | Extended | Remove events |

</td>
<td width="50%" valign="top">

### 📁 **Google Drive** <sub>[`drive_tools.py`](gdrive/drive_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `search_drive_files` | **Core** | Search files with query syntax |
| `get_drive_file_content` | **Core** | Read file content (Office formats) |
| `list_drive_items` | Extended | List folder contents |
| `create_drive_file` | **Core** | Create files or fetch from URLs |

</td>
</tr>
<tr>

<tr>
<td width="50%" valign="top">

### 📧 **Gmail** <sub>[`gmail_tools.py`](gmail/gmail_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `search_gmail_messages` | **Core** | Search with Gmail operators |
| `get_gmail_message_content` | **Core** | Retrieve message content |
| `get_gmail_messages_content_batch` | **Core** | Batch retrieve message content |
| `send_gmail_message` | **Core** | Send emails |
| `get_gmail_thread_content` | Extended | Get full thread content |
| `modify_gmail_message_labels` | Extended | Modify message labels |
| `list_gmail_labels` | Extended | List available labels |
| `manage_gmail_label` | Extended | Create/update/delete labels |
| `draft_gmail_message` | Extended | Create drafts |
| `get_gmail_threads_content_batch` | Complete | Batch retrieve thread content |
| `batch_modify_gmail_message_labels` | Complete | Batch modify labels |
| `start_google_auth` | Complete | Initialize authentication |

</td>
<td width="50%" valign="top">

### 📝 **Google Docs** <sub>[`docs_tools.py`](gdocs/docs_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `get_doc_content` | **Core** | Extract document text |
| `create_doc` | **Core** | Create new documents |
| `modify_doc_text` | **Core** | Modify document text |
| `search_docs` | Extended | Find documents by name |
| `find_and_replace_doc` | Extended | Find and replace text |
| `list_docs_in_folder` | Extended | List docs in folder |
| `insert_doc_elements` | Extended | Add tables, lists, page breaks |
| `insert_doc_image` | Complete | Insert images from Drive/URLs |
| `update_doc_headers_footers` | Complete | Modify headers and footers |
| `batch_update_doc` | Complete | Execute multiple operations |
| `inspect_doc_structure` | Complete | Analyze document structure |
| `create_table_with_data` | Complete | Create data tables |
| `debug_table_structure` | Complete | Debug table issues |
| `*_document_comments` | Complete | Read, Reply, Create, Resolve |

</td>
</tr>

<tr>
<td width="50%" valign="top">

### 📊 **Google Sheets** <sub>[`sheets_tools.py`](gsheets/sheets_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `read_sheet_values` | **Core** | Read cell ranges |
| `modify_sheet_values` | **Core** | Write/update/clear cells |
| `create_spreadsheet` | **Core** | Create new spreadsheets |
| `list_spreadsheets` | Extended | List accessible spreadsheets |
| `get_spreadsheet_info` | Extended | Get spreadsheet metadata |
| `create_sheet` | Complete | Add sheets to existing files |
| `add_chart` | Extended | Insert a native chart (BAR/COLUMN/LINE/AREA/PIE/DOUGHNUT/...) on a sheet from an existing data range, returning its `chart_id`. |
| `*_sheet_comment` | Complete | Read/create/reply/resolve comments |

</td>
<td width="50%" valign="top">

### 🖼️ **Google Slides** <sub>[`slides_tools.py`](gslides/slides_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `create_presentation` | **Core** | Create new presentations |
| `create_audit_presentation` | **Core** | Build a full branded deck from one structured JSON payload (template + tables + images + native Sheets charts + speaker notes). [Schema & example below.](#create_audit_presentation-build-a-full-deck-from-structured-json) |
| `get_presentation` | **Core** | Retrieve presentation details |
| `batch_update_presentation` | Extended | Apply multiple updates |
| `get_page` | Extended | Get specific slide information |
| `get_page_thumbnail` | Extended | Generate slide thumbnails |
| `*_presentation_comment` | Complete | Read/create/reply/resolve comments |

</td>
</tr>
<tr>
<td width="50%" valign="top">

### 📝 **Google Forms** <sub>[`forms_tools.py`](gforms/forms_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `create_form` | **Core** | Create new forms |
| `get_form` | **Core** | Retrieve form details & URLs |
| `set_publish_settings` | Complete | Configure form settings |
| `get_form_response` | Complete | Get individual responses |
| `list_form_responses` | Extended | List all responses with pagination |

</td>
<td width="50%" valign="top">

### ✓ **Google Tasks** <sub>[`tasks_tools.py`](gtasks/tasks_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `list_tasks` | **Core** | List tasks with filtering |
| `get_task` | **Core** | Retrieve task details |
| `create_task` | **Core** | Create tasks with hierarchy |
| `update_task` | **Core** | Modify task properties |
| `delete_task` | Extended | Remove tasks |
| `move_task` | Complete | Reposition tasks |
| `clear_completed_tasks` | Complete | Hide completed tasks |
| `*_task_list` | Complete | List/get/create/update/delete task lists |

</td>
</tr>
<tr>
<td width="50%" valign="top">

### 💬 **Google Chat** <sub>[`chat_tools.py`](gchat/chat_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `list_spaces` | Extended | List chat spaces/rooms |
| `get_messages` | **Core** | Retrieve space messages |
| `send_message` | **Core** | Send messages to spaces |
| `search_messages` | **Core** | Search across chat history |

</td>
<td width="50%" valign="top">

### 🔍 **Google Custom Search** <sub>[`search_tools.py`](gsearch/search_tools.py)</sub>

| Tool | Tier | Description |
|------|------|-------------|
| `search_custom` | **Core** | Perform web searches |
| `get_search_engine_info` | Complete | Retrieve search engine metadata |
| `search_custom_siterestrict` | Extended | Search within specific domains |

</td>
</tr>
</table>


**Tool Tier Legend:**
- <span style="color:#2d5b69">•</span> **Core**: Essential tools for basic functionality • Minimal API usage • Getting started
- <span style="color:#72898f">•</span> **Extended**: Core tools + additional features • Regular usage • Expanded capabilities
- <span style="color:#adbcbc">•</span> **Complete**: All available tools including advanced features • Power users • Full API access

---

### `create_audit_presentation`: build a full deck from structured JSON

Designed for workflows (n8n, custom scripts, agents) that already produce structured audit data
and need to ship it as a branded Google Slides deck without writing any Slides API code.

> 💡 **Driving this from an LLM agent?** A copy-pasteable system prompt that locks the agent to the canonical webloom layout vocabulary and JSON conventions lives at [`gslides/AGENT_SYSTEM_PROMPT.md`](gslides/AGENT_SYSTEM_PROMPT.md). Paste it into Claude / OpenAI / Gemini / n8n's *System Message* and the agent will reliably emit valid `create_audit_presentation` payloads.

One MCP call → whole deck. No iteration. The tool handles template copying, hidden data sheet
creation, native Sheets chart generation, slide layout, speaker notes, folder placement, and
rollback on failure. Slides API requests are chunked internally to stay under per-batch limits;
50 slides take ~30–60 seconds end-to-end.

#### Template contract

You provide the `template_presentation_id` of any Google Slides file. The tool calls
`drive.files.copy()` so the new deck inherits the template's master, layouts, theme colors and
fonts. The template is **never modified**.

- Use Google's predefined layouts (`TITLE`, `TITLE_AND_BODY`, `BLANK`, `SECTION_HEADER`,
  `BIG_NUMBER`, `TITLE_AND_TWO_COLUMNS`, ...) or custom layouts referenced by display name.
- No `{{placeholders}}` needed in the template — content is generated programmatically.
- Required scope: `https://www.googleapis.com/auth/drive` (to copy a user-supplied template that
  the app didn't originally create). The tool requests this scope automatically; expect a
  re-consent screen the first time it runs.

#### `deck` JSON schema

```json
{
  "title": "Pré-audit SEO - edaa.fr - 2026-04",
  "chart_defaults": {
    "series_colors": ["#1A73E8", "#34A853", "#FBBC04", "#EA4335"],
    "background_color": "#FFFFFF",
    "font_family": "Roboto",
    "title_text_format": { "bold": true, "font_size": 14, "foreground_color": "#202124" },
    "legend_position": "BOTTOM_LEGEND"
  },
  "slides": [
    {
      "layout": "Cover",
      "image_placeholders": [
        "https://drive.google.com/uc?export=view&id=LAPTOP_FILE_ID"
      ],
      "speaker_notes": "Slide d'ouverture. Visuel intégré au layout."
    },
    {
      "layout": "Section",
      "fields": { "title": "Pré-audit SEO — edaa.fr", "subtitle": "Avril 2026 · webloom" }
    },
    {
      "layout": "Title + Body",
      "fields": {
        "title": "Synthèse en 30 secondes",
        "body": "📊 edaa.fr affiche un **score SEO global de 59/100**.\n\n📉 Trafic : **8 847 sessions/mois** (-36,1 % YoY).\n\n🚀 4 chantiers prioritaires → **+25 % de trafic en 12 mois** pour **20 000 €** investis."
      },
      "speaker_notes": "Slide la plus importante du deck."
    },
    {
      "layout": "Two Columns",
      "fields": {
        "title": "Avant / Après — 12 mois",
        "body": [
          "📍 **AVANT** (T0)\n\n• Score : **59/100**\n• Trafic : **8 847** /mois\n• CA : **176 940 €** /mois",
          "🎯 **APRÈS** (T+12)\n\n• Score : **80/100**\n• Trafic : **11 000** /mois\n• CA : **220 000 €** /mois"
        ]
      },
      "styles": {
        "body": [null, { "fontFamily": "Inter", "fontSize": { "magnitude": 12, "unit": "PT" } }]
      }
    },
    {
      "layout": "Title + Chart + Body",
      "fields": {
        "title": "Scores SEO par pilier",
        "body": "📖 **Lecture**\n\n• 🛠️ Technique : **90/100**\n• ✍️ Contenu : **64/100**\n• 🔗 Backlinks : **57/100**\n• 🤖 GEO/IA : **10/100**"
      },
      "styles": {
        "body": { "fontFamily": "Inter", "fontSize": { "magnitude": 12, "unit": "PT" } }
      },
      "chart": {
        "type": "COLUMN",
        "title": "Scores /100 par pilier",
        "value_axis_title": "Score",
        "data": {
          "headers": ["Pilier", "Score"],
          "rows": [["Technique", 90], ["Contenu", 64], ["Backlinks", 57], ["GEO / IA", 10]]
        },
        "series_colors": ["#1DB954"],
        "position": { "x": 380, "y": 110, "w": 300, "h": 250 }
      }
    },
    {
      "layout": "Title + Table",
      "fields": { "title": "KPIs principaux" },
      "table": {
        "headers": ["Métrique", "Actuelle", "Cible 12m", "Δ"],
        "rows": [
          ["Score SEO /100", "59", "80", "+36 %"],
          ["Trafic mensuel", "8 847", "11 000", "+24 %"],
          ["CA mensuel (€)", "176 940", "220 000", "+24 %"]
        ],
        "position": { "x": 40, "y": 95, "w": 640, "h": 280 }
      }
    },
    {
      "layout": "Section",
      "fields": { "title": "Merci.", "subtitle": "francois@webloom.fr" }
    }
  ]
}
```

**Slide types** (combine fields freely on a single slide):

| Field | Purpose |
|---|---|
| `layout` | Predefined layout name (`TITLE`, `TITLE_AND_BODY`, `BLANK`, `SECTION_HEADER`, `BIG_NUMBER`, ...) or a custom template layout's display name. |
| `fields.title` / `fields.subtitle` | Fills the single TITLE/SUBTITLE placeholder. |
| `fields.body` | Fills the BODY placeholder. Pass a **string** for a single-body layout, or a **list of strings** for layouts that expose multiple BODY placeholders (e.g. two-column layouts). Item `i` fills the BODY at layout index `i`. Inline `**bold**` markers are parsed and rendered as bold runs; emojis pass through as-is (use `\**` to escape a literal `**`). |
| `styles` | Per-field text style overrides. Each value is a Slides API `TextStyle` dict (`{fontFamily, fontSize:{magnitude,unit}, bold, italic, foregroundColor, ...}`). For two-column layouts, `styles.body` may be an **array** aligned with the body texts; `null` keeps the placeholder default for that slot. **Always set `styles.body[1]` on Two Columns layouts** — the right column renders as an overlay and won't inherit the master style otherwise. |
| `image_placeholders` | List of items targeting **PICTURE placeholders** ("espace réservé image") in the layout. Each item is either a URL string or `{"url": "...", "method": "CENTER_INSIDE"\|"CENTER_CROP"}`. Item `i` fills the PICTURE at layout index `i`. PICTURE placeholders defined as either `shape` or `image` page elements are both detected. |
| `title` (top-level, on `BLANK`) | Adds a free-floating title text box. |
| `table` | `{headers, rows, position?, header_style?, body_style?}` — creates a real `Table` element you can re-style by hand later. |
| `chart` | `{type, title?, data:{headers, rows}, position?, value_axis_title?, domain_axis_title?, legend_position?, series_colors?, stacked_type?, series_types?}` — becomes a native Sheets chart embedded as `LINKED`, so refreshing the Sheet refreshes the deck. |
| `image` | `{url, position?}` — free-floating image, **not** a placeholder. Must be a publicly accessible URL. Use this when the layout has no PICTURE placeholder, or when you want pixel control. |
| `text_boxes` | List of `{text, position?, style?, alignment?}` for free placement. `alignment` ∈ `START` \| `CENTER` \| `END` \| `JUSTIFIED`. |
| `speaker_notes` | Plain text added to the slide's notes page (no markdown). |

**Supported chart types**: `BAR`, `COLUMN`, `LINE`, `AREA`, `SCATTER`, `COMBO`, `STEPPED_AREA`, `PIE`, `DOUGHNUT`. Convention: column 0 of `data.rows` is the X axis (or pie domain); remaining columns are series. Numbers should be numeric (not quoted strings) so Sheets axes auto-format.

Coordinates use **points (PT)**. The default page is 720 × 405 PT (standard widescreen).

**Chart positioning recipes** (default 720×405 slide, title strip occupies y:0–90):

| Layout context | Recommended `position` |
|---|---|
| `Title + Chart + Body` (body on left) | `{ "x": 380, "y": 110, "w": 300, "h": 250 }` (chart on right) |
| `Title + Chart + Body` (body on right) | `{ "x": 40, "y": 110, "w": 300, "h": 250 }` (chart on left) |
| `Title + Chart` (chart only, full width) | `{ "x": 60, "y": 110, "w": 600, "h": 270 }` |

#### Chart styling (brand colors, fonts, background)

Embedded Sheets charts do **not** automatically inherit your Slides template's theme — they are rendered by Sheets, which has its own default palette. To get on-brand charts:

1. **Set deck-wide defaults once** in `deck.chart_defaults`:

   ```json
   "chart_defaults": {
     "series_colors": ["#1A73E8", "#34A853", "#FBBC04", "#EA4335"],
     "background_color": "#FFFFFF",
     "font_family": "Roboto",
     "title_text_format": { "bold": true, "font_size": 14, "foreground_color": "#202124" },
     "legend_position": "BOTTOM_LEGEND",
     "stacked_type": "STACKED"
   }
   ```

2. **Override per chart** by repeating any of those fields inside a single `chart` block. Per-chart values win:

   ```json
   { "type": "COLUMN", "title": "Scores",
     "data": { "headers": ["Pilier", "Score"], "rows": [["Tech", 90], ["Contenu", 64]] },
     "series_colors": ["#FF6F00"],
     "title_text_format": { "bold": true, "font_size": 18, "foreground_color": "#FF6F00" }
   }
   ```

| Style field | Applies to | Notes |
|---|---|---|
| `series_colors` | Bar / Column / Line / Area / Scatter / Combo / Stepped area | List of HEX strings, cycled across series. Pie/doughnut slice colors are not controllable here (Google API limitation). |
| `background_color` | All chart types | HEX, e.g. `"#FFFFFF"`. |
| `font_family` | All chart types | Default font for the whole chart. |
| `title_text_format` | All chart types | `{bold, italic, font_size, font_family, foreground_color}` — all keys optional. |
| `legend_position` | All chart types | `BOTTOM_LEGEND` (default), `LEFT_LEGEND`, `RIGHT_LEGEND`, `TOP_LEGEND`, `NO_LEGEND`. |
| `stacked_type` | Bar / Column / Area only | `NONE`, `STACKED`, `PERCENT_STACKED`. Ignored for other chart types. |

#### Authoring tips & gotchas

These are battle-tested rules captured from real-world deck builds. They apply to whether you call the tool from a chatbot, an n8n workflow, or by hand.

**1. Bold and emojis in `body` text.** Wrap any segment with `**…**` to render it bold; emojis (📊 🚀 🎯 …) pass through transparently. UTF-16 indexing is handled internally so bold ranges align correctly around 4-byte glyphs.

```json
"body": "📊 Score : **59/100**.\n\n🚀 Plan : **+25 %** en 12 mois."
```

To insert a literal `**`, escape it as `\**`.

**2. Two Columns layouts: always pass an explicit `styles.body[1]`.** When a layout exposes two BODY placeholders, the second is rendered as a free-floating `TEXT_BOX` overlay (this works around a known Slides API ghost-bug with multi-BODY placeholders). That overlay does **not** inherit the master text style, so without an explicit style the right column visually drifts from the left. Use `null` as the first array element to keep the left column on its placeholder default:

```json
"styles": {
  "body": [null, { "fontFamily": "Inter", "fontSize": { "magnitude": 12, "unit": "PT" } }]
}
```

Use Two Columns sparingly — single-column slides with a strong `body` paragraph are usually more readable. Reserve two columns for genuine before/after, do/don't, or option A/B comparisons.

**3. PICTURE placeholders ("espace réservé image").** `image_placeholders` only works if the *layout* (not the slide instance) actually exposes one or more PICTURE placeholders. To add one in the editor: **Slide → Edit theme → select the layout → Insert → Placeholder → Image → Apply to layout**. The build log prints, for every slide, which placeholders were detected (`Copy layout placeholders: TITLE x1, BODY x1, PICTURE x1`). If `PICTURE` is missing there, fall back to a free-floating `image` block:

```json
{ "layout": "Cover",
  "image": { "url": "https://…", "position": { "x": 25, "y": 70, "w": 350, "h": 230 } } }
```

PICTURE placeholders can be defined as either `shape` or `image` PageElements in the template — both are detected automatically.

**4. Charts must clear their neighbouring placeholders.** Charts are placed by absolute `position` and do **not** auto-flow inside a column. If you put a chart on a `Title + Chart + Body` slide without `position`, it will overlap the body text. Use the recipes in the table above, or drag-and-tune once in the Slides editor and bake the resulting `x/y/w/h` back into the JSON.

**5. Brand your charts.** Embedded Sheets charts inherit Sheets' default palette, not your Slides theme. Set `deck.chart_defaults` once (series colors, font, title style, legend position), then override per chart only when needed. See the [Chart styling](#chart-styling-brand-colors-fonts-background) section above.

**6. Numbers stay numeric.** In `chart.data.rows`, pass numbers as numbers (`90`), not strings (`"90"`), so Sheets formats axes correctly. In `table.rows`, strings are fine — tables are textual.

**7. Speaker notes are plain text.** No markdown, no inline styling. They go straight into the slide's notes page.

**8. Quotas.** The Slides API enforces ~60 write requests per minute per user. The tool retries transient `429`s automatically; if you regenerate decks frequently from the same project, request a quota bump in Google Cloud Console.

**9. Long builds & MCP client timeouts.** Large decks (~30+ slides, especially with charts and tables) take 60–180 seconds to build because Google's `slides.batchUpdate` is rate-limited and `createSlide` against custom layouts must run sequentially to stay reliable. The tool emits live MCP `progress` notifications during the build (Phase A creation, Phase B content, notes pass) so the streaming-HTTP connection stays warm and the client can show a real progress bar instead of timing out and retrying. **If your MCP client retries the call mid-build, it will start a second parallel build that competes for the same per-user write quota and triggers a 429 cascade.** If you control the client, raise its tool-call timeout to at least 300 s for 30+ slide decks (or 600 s for 100-slide decks). Tested clients:

- **Cursor**: bump the `mcp.tools.timeout` setting (default is usually too low for long builds).
- **Claude Desktop**: edit `claude_desktop_config.json` and add a generous `timeout` to the MCP server entry.
- **n8n / cloud agents**: ensure the workflow's HTTP timeout for the MCP step is generous (≥ 300 s).

#### Example call

```json
{
  "user_google_email": "francois@webloom.fr",
  "template_presentation_id": "1AbCdEfGhIjKlMnOpQrStUvWxYz",
  "folder_path": ["CLIENTS", "edaa.fr", "SEO"],
  "create_folders_if_missing": true,
  "if_exists": "create_new",
  "deck": { "title": "Pré-audit SEO - edaa.fr", "slides": [ /* ... */ ] }
}
```

Response (JSON string):

```json
{
  "presentation_id": "1XYZ...",
  "presentation_url": "https://docs.google.com/presentation/d/1XYZ.../edit",
  "slide_count": 15,
  "data_sheet_url": "https://docs.google.com/spreadsheets/d/1ABC.../edit",
  "data_sheet_id": "1ABC...",
  "folder_id": "0B...",
  "folder_path": "CLIENTS > edaa.fr > SEO",
  "title": "Pré-audit SEO - edaa.fr",
  "message": "Created audit presentation 'Pré-audit SEO - edaa.fr' with 15 slide(s) for francois@webloom.fr. Folder: CLIENTS > edaa.fr > SEO. Data sheet with 4 chart(s): https://..."
}
```

#### Other parameters

- `folder_id` / `folder_path` — pick one. `folder_path` walks/creates nested folders (great for `["CLIENTS", "<domain>", "SEO"]`).
- `if_exists` — `"create_new"` (default; appends a UTC timestamp to the title if a duplicate exists in the folder), `"replace"` (deletes the existing same-titled deck first), or `"skip"` (returns the existing deck untouched).
- `cleanup_data_sheet` — set to `true` to delete the auxiliary Sheet after the deck is built. Default `false` so charts remain refreshable.
- `keep_template_slides` — default `false` (the copied template is wiped first, only the generated slides remain). Set to `true` to **preserve every slide already in the template** and append the generated ones AFTER them. Use this when your template ships with fixed boilerplate (cover, methodology, "À propos", legal mentions, …) that should appear in every audit deck. The original template file is never touched either way; this only controls what happens to the *copy*.

#### Limits & sweet spot

- 1–30 slides: ~10–20 s, single Slides batch.
- 30–80 slides: ~30–60 s, automatically chunked.
- 80–150 slides: ~60–120 s, technically fine but consider splitting into multiple decks per audit pillar.
- > 150 slides: split.

On any failure after the template copy, the tool best-effort deletes the partial deck and data sheet so retries don't accumulate orphans.

---

### Connect to Claude Desktop

The server supports two transport modes:

#### Stdio Mode (Default - Recommended for Claude Desktop)

In general, you should use the one-click DXT installer package for Claude Desktop.
If you are unable to for some reason, you can configure it manually via `claude_desktop_config.json`

**Manual Claude Configuration (Alternative)**

<details>
<summary>📝 <b>Claude Desktop JSON Config</b> <sub><sup>← Click for manual setup instructions</sup></sub></summary>

1. Open Claude Desktop Settings → Developer → Edit Config
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

2. Add the server configuration:
```json
{
  "mcpServers": {
    "google_workspace": {
      "command": "uvx",
      "args": ["workspace-mcp"],
      "env": {
        "GOOGLE_OAUTH_CLIENT_ID": "your-client-id",
        "GOOGLE_OAUTH_CLIENT_SECRET": "your-secret",
        "OAUTHLIB_INSECURE_TRANSPORT": "1"
      }
    }
  }
}
```
</details>

### 2. Advanced / Cross-Platform Installation

If you’re developing, deploying to servers, or using another MCP-capable client, keep reading.

#### Instant CLI (uvx)

<details open>
<summary>⚡ <b>Quick Start with uvx</b> <sub><sup>← No installation required!</sup></sub></summary>

```bash
# Requires Python 3.10+ and uvx
# First, set credentials (see Credential Configuration above)
uvx workspace-mcp --tool-tier core  # or --tools gmail drive calendar
```

> **Note**: Configure [OAuth credentials](#credential-configuration) before running. Supports environment variables, `.env` file, or `client_secret.json`.

</details>

### OAuth 2.1 Support (Multi-User Bearer Token Authentication)

The server includes OAuth 2.1 support for bearer token authentication, enabling multi-user session management. **OAuth 2.1 automatically reuses your existing `GOOGLE_OAUTH_CLIENT_ID` and `GOOGLE_OAUTH_CLIENT_SECRET` credentials** - no additional configuration needed!

**When to use OAuth 2.1:**
- Multiple users accessing the same MCP server instance
- Need for bearer token authentication instead of passing user emails
- Building web applications or APIs on top of the MCP server
- Production environments requiring secure session management
- Browser-based clients requiring CORS support

**Enabling OAuth 2.1:**
To enable OAuth 2.1, set the `MCP_ENABLE_OAUTH21` environment variable to `true`.

```bash
# OAuth 2.1 requires HTTP transport mode
export MCP_ENABLE_OAUTH21=true
uv run main.py --transport streamable-http
```

If `MCP_ENABLE_OAUTH21` is not set to `true`, the server will use legacy authentication, which is suitable for clients that do not support OAuth 2.1.

<details>
<summary>🔐 <b>Innovative CORS Proxy Architecture</b> <sub><sup>← Advanced OAuth 2.1 details</sup></sub></summary>

This implementation solves two critical challenges when using Google OAuth in browser environments:

1.  **Dynamic Client Registration**: Google doesn't support OAuth 2.1 dynamic client registration. Our server provides a clever proxy that accepts any client registration request and returns the pre-configured Google OAuth credentials, allowing standards-compliant clients to work seamlessly.

2.  **CORS Issues**: Google's OAuth endpoints don't include CORS headers, blocking browser-based clients. We implement intelligent proxy endpoints that:
   - Proxy authorization server discovery requests through `/auth/discovery/authorization-server/{server}`
   - Proxy token exchange requests through `/oauth2/token`
   - Add proper CORS headers to all responses
   - Maintain security by only proxying to known Google OAuth endpoints

This architecture enables any OAuth 2.1 compliant client to authenticate users through Google, even from browser environments, without requiring changes to the client implementation.

</details>

**MCP Inspector**: No additional configuration needed with desktop OAuth client.

**Claude Code Inspector**: No additional configuration needed with desktop OAuth client.

### VS Code MCP Client Support

<details>
<summary>🆚 <b>VS Code Configuration</b> <sub><sup>← Setup for VS Code MCP extension</sup></sub></summary>

```json
{
    "servers": {
        "google-workspace": {
            "url": "http://localhost:8000/mcp/",
            "type": "http"
        }
    }
}
```
</details>


#### Reverse Proxy Setup

If you're running the MCP server behind a reverse proxy (nginx, Apache, Cloudflare, etc.), you have two configuration options:

**Problem**: When behind a reverse proxy, the server constructs OAuth URLs using internal ports (e.g., `http://localhost:8000`) but external clients need the public URL (e.g., `https://your-domain.com`).

**Solution 1**: Set `WORKSPACE_EXTERNAL_URL` for all OAuth endpoints:
```bash
# This configures all OAuth endpoints to use your external URL
export WORKSPACE_EXTERNAL_URL="https://your-domain.com"
```

**Solution 2**: Set `GOOGLE_OAUTH_REDIRECT_URI` for just the callback:
```bash
# This only overrides the OAuth callback URL
export GOOGLE_OAUTH_REDIRECT_URI="https://your-domain.com/oauth2callback"
```

You also have options for:
| `OAUTH_CUSTOM_REDIRECT_URIS` *(optional)* | Comma-separated list of additional redirect URIs |
| `OAUTH_ALLOWED_ORIGINS` *(optional)* | Comma-separated list of additional CORS origins |

**Important**:
- Use `WORKSPACE_EXTERNAL_URL` when all OAuth endpoints should use the external URL (recommended for reverse proxy setups)
- Use `GOOGLE_OAUTH_REDIRECT_URI` when you only need to override the callback URL
- The redirect URI must exactly match what's configured in your Google Cloud Console
- Your reverse proxy must forward OAuth-related requests (`/oauth2callback`, `/oauth2/*`, `/.well-known/*`) to the MCP server

<details>
<summary>🚀 <b>Advanced uvx Commands</b> <sub><sup>← More startup options</sup></sub></summary>

```bash
# Configure credentials first (see Credential Configuration section)

# Start with specific tools only
uvx workspace-mcp --tools gmail drive calendar tasks

# Start with tool tiers (recommended for most users)
uvx workspace-mcp --tool-tier core      # Essential tools
uvx workspace-mcp --tool-tier extended  # Core + additional features
uvx workspace-mcp --tool-tier complete  # All tools

# Start in HTTP mode for debugging
uvx workspace-mcp --transport streamable-http
```
</details>

*Requires Python 3.10+ and [uvx](https://github.com/astral-sh/uv). The package is available on [PyPI](https://pypi.org/project/workspace-mcp).*

### Development Installation

For development or customization:

```bash
git clone https://github.com/taylorwilsdon/google_workspace_mcp.git
cd google_workspace_mcp
uv run main.py
```

**Development Installation (For Contributors)**:

<details>
<summary>🔧 <b>Developer Setup JSON</b> <sub><sup>← For contributors & customization</sup></sub></summary>

```json
{
  "mcpServers": {
    "google_workspace": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/path/to/repo/google_workspace_mcp",
        "main.py"
      ],
      "env": {
        "GOOGLE_OAUTH_CLIENT_ID": "your-client-id",
        "GOOGLE_OAUTH_CLIENT_SECRET": "your-secret",
        "OAUTHLIB_INSECURE_TRANSPORT": "1"
      }
    }
  }
}
```
</details>

#### HTTP Mode (For debugging or web interfaces)
If you need to use HTTP mode with Claude Desktop:

```json
{
  "mcpServers": {
    "google_workspace": {
      "command": "npx",
      "args": ["mcp-remote", "http://localhost:8000/mcp"]
    }
  }
}
```

*Note: Make sure to start the server with `--transport streamable-http` when using HTTP mode.*

### First-Time Authentication

The server uses **Google Desktop OAuth** for simplified authentication:

- **No redirect URIs needed**: Desktop OAuth clients handle authentication without complex callback URLs
- **Automatic flow**: The server manages the entire OAuth process transparently
- **Transport-agnostic**: Works seamlessly in both stdio and HTTP modes

When calling a tool:
1. Server returns authorization URL
2. Open URL in browser and authorize
3. Google provides an authorization code
4. Paste the code when prompted (or it's handled automatically)
5. Server completes authentication and retries your request

---

## <span style="color:#adbcbc">◆ Development</span>

### <span style="color:#72898f">Project Structure</span>

```
google_workspace_mcp/
├── auth/              # Authentication system with decorators
├── core/              # MCP server and utilities
├── g{service}/        # Service-specific tools
├── main.py            # Server entry point
├── client_secret.json # OAuth credentials (not committed)
└── pyproject.toml     # Dependencies
```

### Adding New Tools

```python
from auth.service_decorator import require_google_service

@require_google_service("drive", "drive_read")  # Service + scope group
async def your_new_tool(service, param1: str, param2: int = 10):
    """Tool description"""
    # service is automatically injected and cached
    result = service.files().list().execute()
    return result  # Return native Python objects
```

### Architecture Highlights

- **Service Caching**: 30-minute TTL reduces authentication overhead
- **Scope Management**: Centralized in `SCOPE_GROUPS` for easy maintenance
- **Error Handling**: Native exceptions instead of manual error construction
- **Multi-Service Support**: `@require_multiple_services()` for complex tools

### Credential Store System

The server includes an abstract credential store API and a default backend for managing Google OAuth
credentials with support for multiple storage backends:

**Features:**
- **Abstract Interface**: `CredentialStore` base class defines standard operations (get, store, delete, list users)
- **Local File Storage**: `LocalDirectoryCredentialStore` implementation stores credentials as JSON files
- **Configurable Storage**: Environment variable `GOOGLE_MCP_CREDENTIALS_DIR` sets storage location
- **Multi-User Support**: Store and manage credentials for multiple Google accounts
- **Automatic Directory Creation**: Storage directory is created automatically if it doesn't exist

**Configuration:**
```bash
# Optional: Set custom credentials directory
export GOOGLE_MCP_CREDENTIALS_DIR="/path/to/credentials"

# Default locations (if GOOGLE_MCP_CREDENTIALS_DIR not set):
# - ~/.google_workspace_mcp/credentials (if home directory accessible)
# - ./.credentials (fallback)
```

**Usage Example:**
```python
from auth.credential_store import get_credential_store

# Get the global credential store instance
store = get_credential_store()

# Store credentials for a user
store.store_credential("user@example.com", credentials)

# Retrieve credentials
creds = store.get_credential("user@example.com")

# List all users with stored credentials
users = store.list_users()
```

The credential store automatically handles credential serialization, expiry parsing, and provides error handling for storage operations.

---

## <span style="color:#adbcbc">⊠ Security</span>

- **Credentials**: Never commit `.env`, `client_secret.json` or the `.credentials/` directory to source control!
- **OAuth Callback**: Uses `http://localhost:8000/oauth2callback` for development (requires `OAUTHLIB_INSECURE_TRANSPORT=1`)
- **Transport-Aware Callbacks**: Stdio mode starts a minimal HTTP server only for OAuth, ensuring callbacks work in all modes
- **Production**: Use HTTPS & OAuth 2.1 and configure accordingly
- **Network Exposure**: Consider authentication when using `mcpo` over networks
- **Scope Minimization**: Tools request only necessary permissions

---

## <span style="color:#adbcbc">◆ Integration with Open WebUI</span>

<details open>
<summary>◆ <b>Open WebUI Integration</b> <sub><sup>← Connect to Open WebUI as tool provider</sup></sub></summary>

<table>
<tr><td width="50%" valign="top">

### ▶ Instant Start (No Config)
```bash
# Set credentials & launch in one command
GOOGLE_OAUTH_CLIENT_ID="your_id" \
GOOGLE_OAUTH_CLIENT_SECRET="your_secret" \
uvx mcpo --port 8000 --api-key "secret" \
-- uvx workspace-mcp
```

</td><td width="50%" valign="top">

### ◆ Manual Configuration
1. Create `config.json`:
```json
{
  "mcpServers": {
    "google_workspace": {
      "type": "streamablehttp",
      "url": "http://localhost:8000/mcp"
    }
  }
}
```

2. Start MCPO:
```bash
mcpo --port 8001 --config config.json
```

</td></tr>
</table>

### ≡ Configure Open WebUI
1. Navigate to **Settings** → **Connections** → **Tools**
2. Click **Add Tool** and enter:
   - **Server URL**: `http://localhost:8001/google_workspace`
   - **API Key**: Your mcpo `--api-key` (if set)
3. Save - Google Workspace tools are now available!

</details>

---

## <span style="color:#adbcbc">≡ License</span>

MIT License - see `LICENSE` file for details.

---

Validations:
[![MCP Badge](https://lobehub.com/badge/mcp/taylorwilsdon-google_workspace_mcp)](https://lobehub.com/mcp/taylorwilsdon-google_workspace_mcp)

[![Verified on MseeP](https://mseep.ai/badge.svg)](https://mseep.ai/app/eebbc4a6-0f8c-41b2-ace8-038e5516dba0)


<div align="center">
<img width="842" alt="Batch Emails" src="https://github.com/user-attachments/assets/0876c789-7bcc-4414-a144-6c3f0aaffc06" />
</div>
