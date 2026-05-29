# TECHMAP — Claude Code Project Guide

## Project overview
Google Apps Script (GAS) add-on for Google Sheets. Manages manufacturing tech cards (техкарты) for cable harness assembly at ООО «ЗЭП ФРАКСИС». Covers template library, operations DB, and an assembly tech card generator.

## Deploy — MANDATORY after every file edit
```powershell
& "c:\Users\Артемий Васильев\Documents\CURSOR_TECHMAP\deploy.ps1"
```
`clasp push --force` is a **full sync** — deletes from GAS any file not present locally. Always deploy after editing .gs or .html files. Never report a task done without deploying first.

## File structure
| File | Role |
|---|---|
| `Code.gs` | Entry point — menu, `onOpen`, sidebar/dialog launchers |
| `Config.gs` | ALL constants — `TECHMAP_APP`, `TECHOPS_DB_APP`, `ASSEMBLY_GEN` |
| `AssemblyGenerator.gs` | Generator: `generateAssemblyTechCards`, structural fill, placeholder map |
| `AssemblyGeneratorDialog.html` | Dialog UI for generator — op selection, chain preview |
| `OperationDatabase.gs` | DB sync from source spreadsheet, snapshot persistence, payload builder |
| `TemplateStore.gs` | Template CRUD — save/insert/delete from `_TC_STORE` and `_TC_LIBRARY` |
| `WorkspaceSidebar.html` | Main sidebar UI — template browser + operations DB browser |
| `Utils.gs` | Shared helpers — `normalizeHeader_`, `normalizeSearch_`, `ChunkCache_` |
| `RangeCopy.gs` | Range copy preserving formulas and merges |
| `ImageHandler.gs` | Image copy/insert to Drive |
| `SaveTemplateDialog.html` | Dialog for saving a selection as template |
| `appsscript.json` | GAS manifest |

## Local-only (NOT in git, NOT pushed to GAS)
- `.clasp.json` — contains `scriptId`, sensitive
- `deploy.ps1` — PowerShell deploy script

## Key conventions

### Structural fill — no placeholders for DB data
Tech card cells are found by **text pattern detection**, not `{{PLACEHOLDER}}` tokens. The user never inserts markers into templates manually. Existing placeholders (`{{INDEX}}`, `{{NAME}}`, `{{WIRE_NAME}}` etc.) are only for assembly metadata filled from the dialog.

Structural detection patterns in `fillTechCardStructurally_`:
- `Комплектующие` → `/комплектующ/i`
- `Полуфабрикат` → `/полуфабрикат|^п\/ф/`
- `Расчетное время` → `/расс?ч[её]?тное\s*врем/i`
- `Допуск` column → `/^допуск$/i`
- `№ Прог.` column → `/№\s*прог/i`
- `Контроль` rows only → `values[r].slice(0,4).some(c => /^контроль$/i.test(c))`

### Result format (ГОСТ)
- After резка: `силикон 30AWG красный 155мм`
- After опрессовка: `силикон 30AWG красный 155мм, обж. терм. SSHL-002T-P0.2`
- After монтаж: `... → JST-XH-4P ст.А`

### DB schema
`schemaVersion` in `Config.gs` — bump when adding fields to `dataHeaders`. Current: **10**.
`cacheKeyPrefix` — bump together with schemaVersion (format: `techmap-techops-db-vN`).
`dataHeaders` — 17 columns: `tabKey, displayText, normalizedSearch, exportJson, sourceSheet, sortKey, extra1…extra11`.

TER record extra columns mapping:
| Column | extra# | Field |
|---|---|---|
| extra1 | terManufacturer | |
| extra2 | terSeries | |
| extra3 | terComponent | |
| extra4 | terType | |
| extra5 | terArticle | |
| extra6 | terLPlus | |
| extra7 | terLMinus | |
| extra8 | terApplicator | |
| extra9 | terCrimpHeight | |
| extra10 | terPullForceMin | |
| extra11 | terPullForceMax | |

### Template naming convention
Templates are named `"CODE | OperationType"` e.g. `"PRS_TERMINAL_auto | Опрессовка терминалов (А)"`.
The CODE part matches `opNumber` in БД.ОП for time data lookup.

### Op types in generator
| opType | Description |
|---|---|
| `cutWire` | Wire cutting — uses all wires combined |
| `prsTermA` | Crimp terminal side A |
| `prsTermB` | Crimp terminal side B |
| `insTermA` | Mount connector side A |
| `insTermB` | Mount connector side B |

## After schema version bump
User must run **"Синхронизировать базу техопераций"** from menu, or open generator (auto-syncs on schemaVersion mismatch).

## Git
- Branch: `main`
- Remote: GitHub
- Co-authored commits are NOT used (history was cleaned)
- Author email: `artemiyfraxis@gmail.com`
