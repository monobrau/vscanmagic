# VScanMagic .NET Rebuild

Cross-platform .NET 8 rebuild of VScanMagic with Client Review Sessions, template exports, and ConnectSecure integration.

## Quick start

```bash
cd src
dotnet run --project VScanMagic.Web
```

Open `http://127.0.0.1:8080` (override with `VSCANMAGIC_PORT` and `VSCANMAGIC_API_BIND`).

**LAN access:** set `VSCANMAGIC_API_BIND=0.0.0.0`, allow TCP 8080 in Windows Firewall, browse `http://<this-pc-ip>:8080`. Use `archive/Start-VScanMagic-Lan.ps1`. If the UI is unstyled or shows a Blazor error bar, check the browser Network tab for failed `.css` or `_framework/blazor.web.js` requests (WebSockets must work on the same port).

## Projects

| Project | Purpose |
|---------|---------|
| `VScanMagic.Core` | Settings, remediation rules, risk scoring |
| `VScanMagic.Data` | XLSX ingest and Top N scoring |
| `VScanMagic.Review` | Review session SQLite store |
| `VScanMagic.Reports` | DOCX, PDF, ticket, email, flat XLSX export |
| `VScanMagic.ConnectSecure` | ConnectSecure API client |
| `VScanMagic.Web` | Blazor UI + REST API |
| `VScanMagic.Tests` | Unit tests |

## Workflow

1. **Ingest** — Upload or specify path to All Vulnerabilities XLSX
2. **Review** — Presenter UI for live client review (status, tasks, revised remediation)
3. **Export** — Editable DOCX, read-only client PDF, ticket notes, email, XLSX

## Configuration

Reads existing JSON from:

- Windows: `%LOCALAPPDATA%\VScanMagic\`
- Linux/macOS: `~/.config/vscanmagic/` or `$XDG_CONFIG_HOME/vscanmagic/`

## API

- `POST /api/review-sessions` — create session from XLSX path
- `GET/PATCH /api/review-sessions/{id}`
- `POST /api/review-sessions/{id}/export/all`
- Legacy: `POST /api/reports/executive-summary`, `pending-epss`, `all-vulnerabilities`

Set `VSCANMAGIC_API_KEY` to require `ApiKey` on legacy report endpoints.

## Publish

```bash
dotnet publish src/VScanMagic.Web/VScanMagic.Web.csproj -c Release -r linux-x64 --self-contained
```

## Tests

```bash
cd src && dotnet test
```
