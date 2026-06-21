# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Agent Instructions

You're working inside the **WAT framework** (Workflows, Agents, Tools). Probabilistic AI (you) handles reasoning and orchestration; deterministic Python scripts handle execution. That separation is what keeps multi-step accuracy high.

## The WAT Architecture

**Layer 1 — Workflows (`workflows/`):** Markdown SOPs that define the objective, required inputs, which tools to run, expected outputs, and edge-case handling.

**Layer 2 — Agent (you):** Read the relevant workflow, run tools in the correct sequence, handle failures, and ask clarifying questions when needed. Don't try to execute tasks directly when a tool exists for them. Example: to scrape a site, read `workflows/scrape_website.md` then run `tools/scrape_single_site.py`.

**Layer 3 — Tools (`tools/`):** Python scripts that do the actual work — API calls, data transforms, file ops. Credentials live in `.env`.

## Ask First — Hard Rules

- **Creating or overwriting a workflow** — don't draft, replace, or delete without being asked.
- **Running a tool that makes paid API calls or consumes credits** — confirm before each run if the outcome is uncertain.
- **Pushing data to cloud services** (Google Sheets, Slides, etc.) when the destination or content wasn't specified.

Everything else — reading files, running tools with free/local APIs, fixing scripts, updating `.tmp/` — proceed without asking.

## How to Operate

1. **Check `tools/` before building anything.** Only create new scripts when nothing exists for the task.
2. **When a tool fails:** read the full trace, fix the script, retest, then update the workflow with what you learned (rate limits, timing quirks, endpoint changes). If the fix requires paid API calls, confirm first.
3. **Keep workflows current.** When you find a better method or hit a recurring issue, update the workflow — subject to the hard rules above.

## Running Tools

```bash
uv run tools/<script_name>.py       # preferred — handles virtualenv automatically
python tools/<script_name>.py       # fallback if uv is unavailable
```

Install dependencies:

```bash
pip install -r requirements.txt
pip install <package>               # per-tool, if needed
```

## Tool Script Conventions

```python
#!/usr/bin/env python3
from dotenv import load_dotenv
import os, sys

load_dotenv()

def main():
    # single clear responsibility
    # print progress to stdout
    # write outputs to .tmp/ or push to cloud
    # exit(1) with a descriptive message on unrecoverable error
    pass

if __name__ == "__main__":
    main()
```

- One script, one job. No shared state between tools.
- Accept inputs via CLI args or environment variables — never hardcode paths or keys.
- Write intermediate outputs to `.tmp/<descriptive_name>.<ext>`.

## Workflow Execution Pattern

Before any multi-step task:
1. Check `workflows/` for a relevant SOP.
2. Identify the required inputs and tools listed in that workflow.
3. Run each tool in sequence, passing outputs as inputs to the next step.
4. If no workflow exists, ask before creating one.

## File Structure

```
.tmp/                         # Temporary outputs — regenerated as needed, disposable
tools/                        # Python scripts for deterministic execution
workflows/                    # Markdown SOPs
.env                          # API keys and env vars
credentials.json, token.json  # Google OAuth (gitignored)
```

Outputs the user needs to act on go to cloud services (Google Sheets, Slides, etc.), not local files.
