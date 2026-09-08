# Codex Adapter

Use this package as a Codex skill by installing the whole `sheetnext-dev/` directory.

## Install From This Repository

From the `SNEditor` directory:

```powershell
Copy-Item -Recurse -Force .\docs\skill\sheetnext-dev "$env:USERPROFILE\.codex\skills\sheetnext-dev"
```

Restart Codex after copying the directory.

## Repository-Local Use

If the skill is not installed globally, ask Codex to read `docs/skill/sheetnext-dev/SKILL.md` first and then open the specific reference files required by the task.

## Prompt

Use the SheetNext Dev skill. Read `SKILL.md`, choose the relevant module from `references/core-api.md`, and check signatures in `references/api/`. Load integration, event, JSON, enum or AI relay references only when the task needs them.
