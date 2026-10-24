# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Product rules (from the repo owner, all communication in Vietnamese)

- **All development happens in `dist/CongCuKiemTraThue.html` only.** The owner's users have no Python installed. The Python/Streamlit version (`app.py`, `pipeline.py`, `config.py`, `tax_audit/*.py`) is **frozen**: don't edit it, and don't keep it in sync. The README still describes the Python version as having the same results; that is no longer true.
- **Don't change features that already work without asking first.** Every request ends with "các nội dung khác, tab khác giữ nguyên". Touch only what was asked. If a change would affect output columns, sheet layouts or logic outside the request, use AskUserQuestion before editing.
- When the HTML changes, update the version line `<p id="appVersion">Phiên bản YYYY.MM.DD — cập nhật ngày DD/MM/YYYY</p>` in the page header. It tells the owner which copy they have open, since they keep several downloaded copies.
- Commit messages are written in Vietnamese. After pushing, send the updated HTML file to the user and report before/after numbers from the verification run.
- Never commit taxpayer data. `.gitignore` excludes `*.xlsx` except `samples/*.xlsx`.

## The HTML app (`dist/CongCuKiemTraThue.html`)

The app is one self-contained file of about 1 MB that runs offline in a browser, with no build step and no server. It contains four `<script>` blocks:

1. Embedded SheetJS (`XLSX` global), used for every Excel read and write. Don't touch it.
2. Config and utilities: the `DEF` defaults (default column letters, invoice-status keywords, VAT-reduction periods), MST/date/amount parsing, `resolveIndex`, VAT policy, the catalog of goods and services not eligible for the VAT cut, `sheetMatrix`, and ZIP reading through `DecompressionStream`.
3. Processing logic, one section per tab, marked with `/* ===== ... ===== */` banners: the bảng kê audit pipeline (`runAudit` → `buildSheets`), XML → Excel, the foreign-contractor declaration summary (tờ khai nhà thầu NN), TT103 rate checks, file merge, and NPT reconciliation (`buildTmsNpt`, `reconcilePL3`, `applyPL1`, `processPL2`, `buildHokd`, `buildUqMap`).
4. UI and controls: tab switching, file inputs, button handlers (`btnRun`, `btnNpt`, ...), and config save/load.

The tabs are `<section id="tab-…">` elements: audit, xml, xmldecl, tt103, merge, npt. Each handler reads the column-config inputs, calls the pure functions from block 3, stores the result (`RESULT`, `NPTRES`, ...), and renders it with `renderTable`. Downloads go through `writeWB`.

Concepts that span several functions:

- **Column references.** `resolveIndex(header, ref)` accepts a 1-based number, an Excel column letter, or a header name (exact match first, then substring). The global `COL_MODE` (`'letter'` or `'name'`) decides whether a 1–3 letter string is read as a column letter. In name mode, a default like `M` substring-matches any header containing "m", which causes silent wrong-column bugs.
- **Auto-detection.** `detectHeaderRow` and `detectColLetter` (which returns a column letter) usually override user config by header name. For the four PL1 tax columns this override happens in both modes. Those columns are passed as 1-based numbers so they resolve correctly in either mode. The NPT tab prints a "🔎 Cột đã dùng" line showing which columns were actually used.
- **Matching keys.** `mstKey` (strips leading zeros), `cleanMST`/`baseMST` (strips the branch suffix `-00x`), and `stripZeros` (invoice numbers). E-invoice (HĐĐT) matching uses base MST + invoice number + date, and falls back to MST + invoice number (`einvDateMismatch`).
- **Months.** `monthIndex` returns `year*12+month`, `fmtMonth` formats a month index, and `monthsSpan` is inclusive. In NPT reconciliation, a blank TMS "Từ tháng" means January and a blank "Đến tháng" means December of the year under examination. That year comes from the PL3 row's own "Từ tháng".
- **PIT year.** `pitProgressiveYear(income, year)`: the 7-bracket schedule applies up to 2025, the 5-bracket schedule from 2026. The year comes from, in order: the manual `p1_year` input, `detectDeclYear` ("Kỳ kê khai" in the declaration workbook), then fallbacks from PL3.
- **Persistence.** `localStorage` key `taxtool_cfg_v1` holds the audit-tab config and has its own reset button. Key `taxtool_cfg2_v1` (`CFG2_IDS`) is shared by the NPT and merge tabs. Any new config input must be added to `CFG2_IDS` to be saved.

## Verifying changes

There is no test suite. The working method:

```bash
# 1) JS syntax check of all script blocks
node -e "const h=require('fs').readFileSync('dist/CongCuKiemTraThue.html','utf8');for(const m of h.matchAll(/<script>([\s\S]*?)<\/script>/g))new Function(m[1]);console.log('parse OK')"

# 2) Headless browser test: install playwright-core in the scratchpad, not in the repo
#    (package.json and node_modules are gitignored). Chromium is pre-installed:
#    chromium.launch({executablePath:'/opt/pw-browsers/chromium-1194/chrome-linux/chrome'})
```

- Load `file:///.../dist/CongCuKiemTraThue.html`. Either `setInputFiles('#f_npt_qt', …)` and click through `page.evaluate(()=>$('btnNpt').click())` (tab buttons can be hidden, so `page.click` times out), or call page functions directly from `page.evaluate`, for example `reconcilePL3(mat,1,cfg,lk,hokdMap)`, `applyPL1(...)`, `runAudit(wb,'BK',null,null,einvWb,getCfg(),null)`. For `runAudit` with synthetic data, set `cfg.main.headerRow=1`.
- Build test workbooks in the page with `XLSX.utils.aoa_to_sheet` and `XLSX.write(...,{type:'base64'})`. openpyxl is often not installed, and files the user uploaded under `/root/.claude/uploads/` disappear when the environment resets.
- **Regression check:** run the same input through `git show HEAD:dist/CongCuKiemTraThue.html` and through the working copy, then compare `JSON.stringify` of the result objects. Only the intended columns or sheets may differ.
- `python scripts/make_samples.py` regenerates `samples/` for the frozen Python version only.
