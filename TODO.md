# TODO — SQL Server → MariaDB Migration

Stand: 2026-04-28
Repo: `github.com/McKloony/Powershell_Datenimport`
Letzter Run: `Migrate-SqlToMariaDB.ps1 -Mode Migrate -SkipBackup` — bricht in Step 7 ab.

## Sofortiger Wiedereinstieg

```cmd
cd C:\Users\schmi\Documents\devcode1
git pull
Migrate-SqlToMariaDB.cmd                    REM Menü
Migrate-SqlToMariaDB.cmd -Mode Verify -SkipBackup
Migrate-SqlToMariaDB.cmd -Mode Migrate -SkipBackup
```

Connection-Strings: siehe `db-zugriff-vm.md` (privat).

## Offene Bugs (aus den Runs vom 2026-04-28)

### Bug 1 — Step 6: `fld_identifier` zu lang

- Symptom: `Data too long for column 'fld_identifier' at row 80` in `tbl_documentation_main`.
- Fundstelle: `Migrate-SqlToMariaDB.ps1` Zeile **1286** in `Step6_Migrate-Documentation`.
  ```powershell
  id = ConvertTo-DbString $row.GONr -MaxLen 10 -Required
  ```
- Hypothese: Zielspalte ist kürzer als 10 Zeichen.
- Fix-Idee: tatsächliche Länge aus `INFORMATION_SCHEMA.COLUMNS` lesen und `MaxLen` daran anpassen
  (oder fix auf den korrekten DDL-Wert setzen).

### Bug 2 — Step 7: `fld_number` Out of Range

- Symptom: `Out of range value for column 'fld_number' at row 144` in `tbl_receipt_content`.
- Fundstelle: `Migrate-SqlToMariaDB.ps1` Zeile **1383** in `Step7_Migrate-ReceiptContent`.
  ```powershell
  n = if ($row.XCnt -is [System.DBNull]) { 1 } else { ConvertTo-SmallInt $row.XCnt }
  ```
- Hypothese: Zielspalte ist TINYINT (UNSIGNED 0..255 oder SIGNED -128..127), nicht SMALLINT.
- Fix-Idee: `ConvertTo-UTinyInt` (oder neuen `ConvertTo-TinyInt`) verwenden, Typ via DDL-Lookup verifizieren.

### Bug 3 — Step 7: 71 % Skip-Quote

- Run 11:50: 75.754 gelesen → 21.450 inserted, **54.154 skipped** (54.154 Warnings).
- Vermutlich greifen Tenant-, Receipt- oder Service-Lookup für die Mehrheit der `Tabelle_Abre`-Zeilen mit `ID1 > 10` nicht.
- Diagnose: in `Step7_Migrate-ReceiptContent` Skip-Zähler **pro Reason** loggen
  (no tenant / no receipt / no service / no entity).

## Weitere Auffälligkeiten (nicht fatal)

- `tbl_contact_main`: 2 Zeilen `fld_medical_history` Truncation (Quell-IDs 3271, 8307) → Spaltenlänge prüfen.
- `tbl_receipt_main`: 10 Warnings — Detail im Transcript, vermutlich Truncation.

## Diagnose-Snippet — MariaDB-Spaltentypen

```sql
SELECT TABLE_NAME, COLUMN_NAME, COLUMN_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = 'simplimed'
  AND TABLE_NAME IN ('tbl_documentation_main','tbl_receipt_content','tbl_contact_main','tbl_receipt_main')
  AND COLUMN_NAME IN ('fld_identifier','fld_number','fld_medical_history')
ORDER BY TABLE_NAME, COLUMN_NAME;
```

## Pending — Stage 2 (`Merge-MariaToMaria.ps1`)

Noch nicht angefangen. Erst soll die SQL→MariaDB-Migration sauber durchlaufen.

Konzept:
- Lesen aus `simplimed_staging`, schreiben in produktive `simplimed`.
- AUTO_INCREMENT-Range + GUID-Match für Idempotenz.
- Tenant-Filter (nur explizit freigegebene Tenants).
- Kein TRUNCATE im Produktionspfad — In-Place-Append.

## Konsolidiertes Skript-Inventar (Stand)

| Datei                       | Rolle                                              |
| --------------------------- | -------------------------------------------------- |
| `Migrate-SqlToMariaDB.ps1`  | Hauptskript, alle Modi (Migrate/Verify/Repair/Cleanup) |
| `Migrate-SqlToMariaDB.cmd`  | Host-Launcher (`powershell.exe -ExecutionPolicy Bypass`) |
| `SQL_Server.sql`            | Quell-DB-Snapshot                                  |
| `table_datafeld_mapping.csv`| Source→Target Feldmapping                          |
| `db-zugriff-vm.md`          | Connection-Infos (privat)                          |
| `Promt.txt`                 | Projekt-Briefing                                   |
