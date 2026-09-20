# Contributing

Thanks for your interest in improving this toolkit.

## What to contribute

- Fixes and features for **modern SPO** scripts under `CreateSiteCollections/` and `Get_SPOInventory/`
- Documentation clarity (especially root and folder READMEs)
- Safer auth patterns (Entra app + certificate), validation, and logging

Historical folders (`MOSS2007`, `SP2010`, `SP2013`, large parts of `SPO`, `Scripts`) are primarily an **archive**. Large refactors there are optional; prefer clear warnings over silent “modernization.”

## Guidelines

1. Open an issue for sizable changes when practical.  
2. Use a feature branch and a focused pull request.  
3. Do **not** commit secrets, production CSVs, certificates, or real tenant identifiers.  
4. Match existing PowerShell style where you edit a script; add comment-based help for new parameters.  
5. Update the relevant README when behavior or paths change.

## Local checks

- Run scripts only against a **dev** tenant.  
- Prefer PowerShell 7.4+ with a current PnP.PowerShell 3.x build for modern scripts.  
- Confirm `.gitignore` still excludes logs and local reports.

## Code of conduct

Be respectful. This is a personal open-source collection maintained in spare time.
