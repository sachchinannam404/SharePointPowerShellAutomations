# Security

## Supported content

Security-relevant attention focuses on **modern SharePoint Online** scripts:

- `CreateSiteCollections/`
- `Get_SPOInventory/`

Archived on-premises and general utility scripts are provided **as-is** for reference. Do not use them against production Microsoft 365 without careful review.

## Reporting a vulnerability

Please **do not** open a public issue for sensitive security reports.

1. Contact the repository owner via GitHub (private security advisory if available, or a private channel linked from the profile).  
2. Include script path, impact, and a minimal reproduction **without** real tenant secrets.

## Safe usage

- Register **your own** Entra ID application for PnP.PowerShell; do not rely on retired multi-tenant helper apps.  
- Prefer **certificate-based app-only** auth for automation; avoid embedding secrets in scripts.  
- Store certificates and client IDs outside the repo (Key Vault, pipeline secrets, user cert store).  
- Use least-privilege SharePoint/Graph permissions and separate dev/test tenants.  
- Never commit `.pfx`, `.pem`, `.env`, or production inventory CSVs (see `.gitignore`).
