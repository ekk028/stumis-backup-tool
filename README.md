# STU-MIS Backup Tool — Handoff Package

This folder is ready to hand off to **Claude Code** to build a Windows desktop tool that backs up data from the STU-MIS school management platform.

## How to use this with Claude Code

### Step 1 — Install Claude Code
- Windows: see `https://docs.claude.com/claude-code` for the latest install command
- Sign in with your Anthropic / Claude account

### Step 2 — Open this folder in Claude Code
```bash
cd "D:\Downloads\stumis-backup-tool-project"
claude
```

### Step 3 — Give Claude Code this prompt to start

```
Read CLAUDE.md, ENDPOINTS.md, and SPEC.md first.

Then build the STU-MIS Backup Tool as described. Start with the MVP
(Cat 2/3/5/6 — bulk Excel endpoints) and a minimal customtkinter UI.
Test it works against https://wangluangpitt.stumis06.com/ then add
the remaining categories one at a time.

The starter_code/ folder has my JSON→Excel converters from manually
backing up one school. Use them as reference for the output format.
sample_data/ has small JSON samples — you can dev-test conversion
without needing to hit the real server every time.

After MVP works, build the .exe with PyInstaller and confirm it runs
on Windows.
```

### Step 4 — Iterate
Claude Code will work in stages. Each session may take 2-4 hours of active work. Expect 3-5 sessions total to ship a polished v1.

## What's in this folder

| File | Purpose |
|---|---|
| `CLAUDE.md` | Operating manual for Claude Code — read first |
| `SPEC.md` | Detailed requirements & user stories |
| `ENDPOINTS.md` | All STU-MIS HTTP endpoints + request formats |
| `starter_code/` | Working Python converters from Cowork session |
| `sample_data/` | Real JSON dumps from one school (truncated) |

## Real-world test credentials
You'll need a real school's login to test. The team that built this used `wangluangpitt.stumis06.com` with admin credentials. Use your own school's URL and admin login for testing.

## Authoring notes
The handoff package was prepared by Claude (in Cowork mode) after completing one school's full backup manually — that's where the endpoint knowledge came from. The Excel converter logic is already correct; the hard work is wrapping it in a clean UI and login flow.
