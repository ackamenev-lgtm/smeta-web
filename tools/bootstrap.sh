#!/bin/bash
# Один раз на машине агента, если ещё нет LibreOffice. В собранном окружении уже стоит.
set -euo pipefail
if ! command -v soffice >/dev/null 2>&1; then
  sudo DEBIAN_FRONTEND=noninteractive apt-get update -qq
  sudo DEBIAN_FRONTEND=noninteractive apt-get install -y -qq \
    libreoffice-calc libreoffice-writer poppler-utils fonts-dejavu-core
fi
python3 -c "import openpyxl, docx" 2>/dev/null || python3 -m pip install --user openpyxl python-docx
echo "готово: $(command -v soffice)"
