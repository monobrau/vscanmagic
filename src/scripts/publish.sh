#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
RID="${1:-linux-x64}"
dotnet publish VScanMagic.Web/VScanMagic.Web.csproj -c Release -r "$RID" --self-contained -o "publish/$RID"
echo "Published to publish/$RID"
