# Agent Rules

## GitHub Operations

Any operation that updates GitHub remotely must receive explicit user approval before execution.

This includes, but is not limited to:

- `git push`
- Creating a GitHub repository
- Updating a remote repository
- Uploading a new `win.exe`
- Creating or updating a GitHub Release
- Deleting, overwriting, or force-updating remote content

Local edits, tests, and builds may be performed without GitHub upload. Before syncing any local change to GitHub, stop and ask the user for confirmation.
