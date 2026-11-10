---
name: "kp-document-pipeline"
description: "Node.js pipeline for generating and verifying commercial-proposal (КП) DOCX/PDF documents — read for tooling, rendering fixes, and standard КП structure"
---

- Builds document generation pipelines for commercial proposals (КП) programmatically in Node.js using the `docx` library
- DOCX→PDF conversion via LibreOffice headless; PDF verification with `pdfinfo`, `pdftotext -layout`, `pdftoppm`, and PIL/numpy
- Uses Pandoc (`pandoc -t markdown`) to parse uploaded .docx source files for content extraction
- A4 landscape rendering fix: always pass portrait A4 dimensions (`width: 11906, height: 16838`) together with `PageOrientation.LANDSCAPE`; passing pre-swapped dimensions with the landscape flag causes the `docx` library to double-swap back to portrait
- Table column widths must be validated against available print width (page width minus margins, in DXA)
- Standard КП structure: company letterhead, addressee block, pricing tables with columns for both без НДС and с НДС 22% (treated as binding prices, not reference figures), conditions section, and a Director signature/stamp block
- ООО «ДАТА ВАЙЗ» operates under УСН and pays НДС at 22%; General Director is Козырев М. В.
- Completed two КП documents for АО «Тандер» (сеть «Магнит») — one for the СПб Восток branch (contact Давыдов И. А.) and one for the СПб Север branch (contact Штельма А. Ю.)
