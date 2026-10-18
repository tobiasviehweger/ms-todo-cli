# ms-todo-cli

A command-line interface for Microsoft To Do, built with TypeScript and Node.js.

## Features

- OAuth authentication via Microsoft device code flow
- Manage todo lists (create, list)
- Manage list groups (create, update, delete, list)
- Manage tasks (create, update, complete, list, get)
- Manage task steps / checklist items (create, update, complete, delete, list)
- JSON output for easy scripting and integration
- Multiple Microsoft accounts with user-defined profiles and isolated token caches
- Private authentication storage under `~/.ms-todo-cli/` (directories 0o700, files 0o600)

## Installation

```bash
npm install -g ms-todo-cli
```

Or clone and build locally:

```bash
git clone https://github.com/your-org/ms-todo-cli
cd ms-todo-cli
npm install
npm run build
npm link
```

## Help

Run `ms-todo-cli --help` (or `-h`) to see available commands. Add `--help` to
any command to see its parameters:

```bash
ms-todo-cli task --help
ms-todo-cli task create --help
```

Running without a command, or with only a command group such as `ms-todo-cli task`,
prints usage help to stderr and exits with code 1. Explicit help and `--version`
print to stdout and exit successfully; neither requires authentication.

## Azure App Registration

To use this CLI, you need to register an application in Azure Active Directory:

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click **New registration**
   - Name: `ms-todo-cli` (or any name)
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: leave blank (device code flow doesn't need one)
3. Click **Register**
4. Copy the **Application (client) ID**
5. Under **API permissions**, add:
   - `Tasks.ReadWrite` (Microsoft Graph, Delegated)
   - `User.Read` (Microsoft Graph, Delegated)
   - `offline_access` (Microsoft Graph, Delegated)
6. Grant admin consent if required by your organization

## Environment Variables

| Variable | Description | Required |
|---|---|---|
| `MS_TODO_CLIENT_ID` | Azure app client ID | Required for login and Graph operations, not cached account/status/profile listing or logout |
| `MS_TODO_PROFILE` | Account profile when `--profile` is omitted; otherwise `default` | No |

Set your client ID before using:

```bash
export MS_TODO_CLIENT_ID="your-azure-client-id"
```

## Authentication

### Login

```bash
ms-todo-cli auth login
```

Follow the device code instructions printed to stderr. After login, a success JSON is printed to stdout.

### Multiple accounts / profiles

Choose your own aliases when signing in:

```bash
ms-todo-cli auth login --profile personal
ms-todo-cli auth login --profile work
ms-todo-cli auth profiles
```

Aliases are 1–64 lowercase letters, digits, hyphens or underscores, starting with a
letter or digit. Each profile is bound to the Microsoft account identity returned
at login, rather than its email address or its position in the token cache.
Its account metadata and MSAL cache are stored together in
`~/.ms-todo-cli/profiles/<alias>.json`, atomically updated under a per-profile
filesystem lock so concurrent CLI processes do not overwrite token refreshes.
Signing into a different account under an existing alias requires explicit consent:

```bash
ms-todo-cli auth login --profile work --replace
```

Every command, including authentication, attachments, steps, groups and the
top-level shortcuts, accepts `--profile` before, between or after subcommands:

```bash
ms-todo-cli --profile work task list --all-lists
ms-todo-cli task --profile personal create --title "Buy milk"
ms-todo-cli create --title "Review proposal" --profile work
```

Selection precedence is **`--profile` → `MS_TODO_PROFILE` → `default`**.
There is no mutable active account. An unknown or expired profile never falls back
to another account, and task JSON read from stdin cannot change the profile.
Task and list response structures remain unchanged.

Authentication status reports whether the selected profile has a cached account;
it includes `tokenVerified: false` and does not verify that its token is usable.
`auth profiles` returns aliases and cached account details without token data.
Normal Graph commands refresh tokens
silently when possible and return a profile-specific reauthentication error when
interactive login is needed. Only `auth login` starts device-code authentication.
Logout removes authentication only for the selected profile:

```bash
ms-todo-cli auth status --profile work
ms-todo-cli auth print-account --profile work
ms-todo-cli auth logout --profile work
```

Existing single-account installations retain the `default` profile. The legacy
`~/.ms-todo-cli/msal-cache.json` cache is migrated safely when unambiguous; a cache
containing multiple accounts is rejected rather than choosing the first account.
All profiles share `MS_TODO_CLIENT_ID` and the existing Microsoft Graph permissions.
Profile credentials are local sensitive files: do not commit or share them.

### Check status

```bash
ms-todo-cli auth status
```

### Print account info

```bash
ms-todo-cli auth print-account
```

### Logout

```bash
ms-todo-cli auth logout
```

## List Commands

### Create a list

```bash
ms-todo-cli list create "Shopping"
```

### List all lists

```bash
ms-todo-cli list list
```

## List Group Commands

### Create a list group

```bash
ms-todo-cli group create "Work"
```

### List all list groups

```bash
ms-todo-cli group list
```

### Update a list group name

```bash
ms-todo-cli group update --group-id "GROUP_ID" --name "Renamed"
# Or resolve by name:
ms-todo-cli group update --group "Work" --name "Projects"
```

### Delete a list group

```bash
ms-todo-cli group delete --group-id "GROUP_ID"
# Or resolve by name:
ms-todo-cli group delete --group "Work"
```

## Task Commands

### Create a task

```bash
ms-todo-cli task create --title "Buy milk" --list "Shopping"
ms-todo-cli task create --title "Buy milk" --due "2024-12-31T00:00:00Z" --priority high
```

### Create a task via stdin (JSON)

```bash
echo '{"title":"Buy eggs","list":"Shopping"}' | ms-todo-cli task create --stdin
```

### Update a task

```bash
ms-todo-cli task update --task-id "TASK_ID" --title "Buy oat milk"
ms-todo-cli task update --task-id "TASK_ID" --completed true
# Provide --list-id to skip the O(N) list scan:
ms-todo-cli task update --task-id "TASK_ID" --list-id "LIST_ID" --title "Buy oat milk"
```

### Complete a task

```bash
ms-todo-cli task complete --task-id "TASK_ID"
# Provide --list-id to skip the O(N) list scan:
ms-todo-cli task complete --task-id "TASK_ID" --list-id "LIST_ID"
```

### List tasks in a list

```bash
ms-todo-cli task list --list "Shopping"
ms-todo-cli task list --list-id "LIST_ID"
```

### List tasks across all lists

```bash
ms-todo-cli task list --all-lists
```

### Get a single task

```bash
ms-todo-cli task get --task-id "TASK_ID"
# Provide --list-id to skip the O(N) list scan:
ms-todo-cli task get --task-id "TASK_ID" --list-id "LIST_ID"
```

### Attach a file to a task

```bash
ms-todo-cli task attach --task-id "TASK_ID" --file ./notes.txt
# Provide --list-id to skip the O(N) list scan and --name to override the uploaded filename:
ms-todo-cli task attach --task-id "TASK_ID" --list-id "LIST_ID" --file ./notes.txt --name task-notes.txt
```

Attachments use Microsoft Graph simple attachments (max 3 MiB / 3,145,728 bytes per file).

### Search tasks across all lists by keyword

```bash
ms-todo-cli task search --query "keyword"
# Positional keyword alias:
ms-todo-cli task search "keyword"
```

## Task Step Commands

Microsoft To Do supports checklist items (steps) within a task.

### List steps for a task

```bash
ms-todo-cli step list --task-id "TASK_ID"
```

### Create a step

```bash
ms-todo-cli step create --task-id "TASK_ID" --title "Sub-task title"
```

### Update a step title

```bash
ms-todo-cli step update --task-id "TASK_ID" --step-id "STEP_ID" --title "New title"
```

### Complete a step

```bash
ms-todo-cli step complete --task-id "TASK_ID" --step-id "STEP_ID"
```

### Delete a step

```bash
ms-todo-cli step delete --task-id "TASK_ID" --step-id "STEP_ID"
```

## Short-form Aliases

Top-level shortcuts for common task operations:

```bash
ms-todo-cli create --title "Buy milk"
ms-todo-cli update --task-id "TASK_ID" --title "Buy oat milk"
ms-todo-cli get --task-id "TASK_ID"
```

## Output Format

Command responses are valid JSON; help and version output are plain text.
Success responses have `ok: true`:

```json
{"ok":true,"task":{"id":"...","title":"Buy milk","status":"notStarted"}}
```

Error responses have `ok: false` and exit with code 1:

```json
{"ok":false,"error":{"code":"AUTH_REQUIRED","message":"Not authenticated. Run: ms-todo-cli auth login"}}
```

## OpenClaw / Scripting Integration

The JSON output is designed for easy parsing:

Set `MS_TODO_CLIENT_ID` and optionally `MS_TODO_PROFILE` in the OpenClaw process
environment. Complete device-code login once for each alias using the same OS user
and home directory as OpenClaw (or its persistent container home). Subsequent calls
use cached credentials without prompting. Keep the authentication directory on a
private persistent volume if the environment is recreated.

For concurrent jobs, pass the profile explicitly on each invocation, or set it
only in that job's environment; do not modify a shared active-account setting.

```bash
MS_TODO_PROFILE=work ms-todo-cli task list --all-lists
ms-todo-cli task search --query "meeting" --profile work
echo '{"title":"Buy eggs","list":"Shopping"}' | ms-todo-cli create --stdin --profile personal
```

```bash
# Get all lists and parse with jq
ms-todo-cli list list | jq '.lists[].displayName'

# Create a task and capture its ID
TASK_ID=$(ms-todo-cli create --title "Review PR" | jq -r '.task.id')

# Complete the task
ms-todo-cli task complete --task-id "$TASK_ID"
```

## Development

```bash
npm run build   # Compile TypeScript
npm test        # Run tests
npm run lint    # Lint source files
```

## License

MIT
