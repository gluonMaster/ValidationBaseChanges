"""
Management command: import_access_year

Imports data from an Access database (.accdb) into Django for a specific year.

Usage:
    python manage.py import_access_year --year 2025 --access-file KindElternDaten_25_front.accdb
    python manage.py import_access_year --year 2025 --access-file /path/to/file.accdb --dry-run
    python manage.py import_access_year --year 2024 --access-file data.accdb --familyid-policy=auto-merge
    python manage.py import_access_year --year 2025 --access-file data.accdb --sync-history
    python manage.py import_access_year --year 2025 --access-file data.accdb --patch-fields --dry-run

Options:
    --year            (required) Year for the imported records
    --access-file     (required) Access database file name or path
    --dry-run         Analyze only, don't write to database
    --patch-fields    Patch mode: update only teacher names, contract type/status,
                      and sepa_marker for existing records (no full re-import)
    --familyid-policy Policy for handling FamilyID issues: 'report' (default) or 'auto-merge'
    --sync-history    After import, sync history_raw to HistoryEvent models
    --skip-pending    Skip importing pre_tblKartei (pending changes)
    --skip-declined   Skip importing decl_tblKartei (declined changes)
    --report-dir      Directory for CSV/JSON reports (default: current dir)

Patch mode (--patch-fields):
    Updates ONLY these fields for existing records (by year, id):
    - teacher1_legacy_name (from Value11)
    - teacher2_legacy_name (from Value16)
    - contract_type_raw (from Value14) + computed is_monthly_contract
    - contract_status_raw (from Value20) + computed is_contract_terminated
    - sepa_marker (from Value47)

See PROMPT_07_LEGACY_IMPORT.md for detailed documentation.
"""

from __future__ import annotations

import csv
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

from django.core.management.base import BaseCommand, CommandError
from django.db import connection

if TYPE_CHECKING:
    from apps.legacy_import.services import FamilyIdIssue, ImportStats

logger = logging.getLogger(__name__)


class Command(BaseCommand):
    help = "Import data from Access database (.accdb) into Django for a specific year"

    def _ensure_utf8_stdio(self) -> None:
        """
        Best-effort: ensure we can print Unicode (German UI texts) without crashing
        on Windows consoles with legacy code pages.
        """
        for stream in (sys.stdout, sys.stderr):
            try:
                if hasattr(stream, "reconfigure"):
                    stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                continue

    def _ensure_patch_schema_ready(self, required_columns: set[str]) -> None:
        from apps.karteien.models import KarteiRecord

        table_name = KarteiRecord._meta.db_table
        with connection.cursor() as cursor:
            try:
                description = connection.introspection.get_table_description(cursor, table_name)
            except Exception as exc:
                raise CommandError(
                    "Patch-Import kann nicht starten: Tabellenstruktur kann nicht gelesen werden.\n"
                    f"Tabelle: {table_name}\n"
                    f"Fehler: {exc}"
                ) from exc

        existing_columns = {getattr(col, "name", col[0]) for col in description}
        missing = sorted(required_columns - existing_columns)
        if not missing:
            return

        missing_list = ", ".join(missing)
        raise CommandError(
            "Patch-Import kann nicht starten, weil die Datenbank-Schema veraltet ist.\n"
            f"Fehlende Spalten in Tabelle '{table_name}': {missing_list}\n\n"
            "Bitte Migrationen gegen dieselbe Datenbank ausführen, z.B.:\n"
            "  python manage.py migrate\n"
            "oder (nur App karteien):\n"
            "  python manage.py migrate karteien\n\n"
            "Wenn du Docker verwendest (DB im Container):\n"
            "  docker compose exec web python manage.py migrate"
        )

    def add_arguments(self, parser):
        parser.add_argument(
            "--year",
            type=int,
            required=True,
            help="Year for the imported records (e.g., 2025)",
        )
        parser.add_argument(
            "--access-file",
            type=str,
            required=True,
            help="Access database file name or full path",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            default=False,
            help="Analyze only, don't write to database",
        )
        parser.add_argument(
            "--patch-fields",
            action="store_true",
            default=False,
            help="Patch mode: update only teacher names, contract type/status, and sepa_marker "
                 "for existing records (does not import new records or modify other fields)",
        )
        parser.add_argument(
            "--familyid-policy",
            type=str,
            choices=["report", "auto-merge"],
            default="report",
            help="Policy for handling FamilyID issues (default: report)",
        )
        parser.add_argument(
            "--sync-history",
            action="store_true",
            default=False,
            help="Sync history_raw to HistoryEvent models after import",
        )
        parser.add_argument(
            "--skip-pending",
            action="store_true",
            default=False,
            help="Skip importing pre_tblKartei (pending changes)",
        )
        parser.add_argument(
            "--skip-declined",
            action="store_true",
            default=False,
            help="Skip importing decl_tblKartei (declined changes)",
        )
        parser.add_argument(
            "--report-dir",
            type=str,
            default=".",
            help="Directory for CSV/JSON reports",
        )

    def handle(self, *args, **options):
        self._ensure_utf8_stdio()

        year = options["year"]
        access_file = options["access_file"]
        dry_run = options["dry_run"]
        patch_fields = options["patch_fields"]
        familyid_policy = options["familyid_policy"]
        sync_history = options["sync_history"]
        skip_pending = options["skip_pending"]
        skip_declined = options["skip_declined"]
        report_dir = Path(options["report_dir"])

        # Patch mode is a separate workflow
        if patch_fields:
            return self._handle_patch_mode(year, access_file, dry_run, report_dir)

        # Ensure report directory exists
        report_dir.mkdir(parents=True, exist_ok=True)

        self.stdout.write(self.style.NOTICE(f"\n{'='*60}"))
        self.stdout.write(self.style.NOTICE("Access → Django Import"))
        self.stdout.write(self.style.NOTICE(f"{'='*60}"))
        self.stdout.write(f"Year: {year}")
        self.stdout.write(f"Access file: {access_file}")
        self.stdout.write(f"Dry run: {dry_run}")
        self.stdout.write(f"FamilyID policy: {familyid_policy}")
        self.stdout.write(f"Sync history: {sync_history}")
        self.stdout.write(f"Report directory: {report_dir}")
        self.stdout.write("")

        if dry_run:
            self.stdout.write(self.style.WARNING("DRY RUN MODE - No changes will be written to the database\n"))

        # Import modules here to avoid issues if pyodbc is not installed
        try:
            from apps.legacy_import.access_client import (
                resolve_access_file,
                open_access_connection,
                load_tbl_kartei,
                load_pre_tbl_kartei,
                load_decl_tbl_kartei,
                get_database_stats,
            )
            from apps.legacy_import.services import (
                ImportStats,
                FamilyIdIssue,
                import_tbl_kartei,
                import_pre_tbl_kartei,
                import_decl_tbl_kartei,
                sync_history_for_records,
                analyze_familyid_issues,
            )
        except ImportError as e:
            raise CommandError(f"Failed to import required modules: {e}")

        # Resolve Access file path
        try:
            file_path = resolve_access_file(access_file)
        except FileNotFoundError as e:
            raise CommandError(str(e))
        except ValueError as e:
            raise CommandError(str(e))

        self.stdout.write(f"Resolved file path: {file_path}")

        # Open connection and import
        try:
            with open_access_connection(file_path) as conn:
                # Get database stats
                self.stdout.write(self.style.NOTICE("\nDatabase Statistics:"))
                stats = get_database_stats(conn, year)
                for table_name, count in stats.items():
                    self.stdout.write(f"  {table_name}: {count} rows")

                # Load all data into memory for analysis
                self.stdout.write(self.style.NOTICE("\nLoading data from Access..."))

                tbl_kartei_rows = list(load_tbl_kartei(conn, year))
                self.stdout.write(f"  Loaded {len(tbl_kartei_rows)} rows from tblKartei")

                pre_rows = []
                if not skip_pending:
                    pre_rows = list(load_pre_tbl_kartei(conn, year))
                    self.stdout.write(f"  Loaded {len(pre_rows)} rows from pre_tblKartei")

                decl_rows = []
                if not skip_declined:
                    decl_rows = list(load_decl_tbl_kartei(conn, year))
                    self.stdout.write(f"  Loaded {len(decl_rows)} rows from decl_tblKartei")

        except ImportError as e:
            raise CommandError(
                f"pyodbc is not installed. Install it with: pip install pyodbc\n"
                f"Error: {e}"
            )
        except Exception as e:
            raise CommandError(f"Failed to connect to Access database: {e}")

        # Analyze FamilyID issues
        self.stdout.write(self.style.NOTICE("\nAnalyzing FamilyID consistency..."))

        # Get existing records for comparison
        from apps.karteien.models import KarteiRecord
        existing_records = {}
        for rec in KarteiRecord.objects.filter(year=year).values(
            "id", "family_id", "parent_name", "email"
        ):
            existing_records[(year, rec["id"])] = rec

        all_rows = tbl_kartei_rows + pre_rows + decl_rows
        familyid_issues, auto_merge_map = analyze_familyid_issues(
            all_rows, existing_records, year
        )

        if familyid_issues:
            self.stdout.write(
                self.style.WARNING(f"  Found {len(familyid_issues)} FamilyID issues")
            )
            self._write_familyid_report(familyid_issues, year, report_dir)
        else:
            self.stdout.write(self.style.SUCCESS("  No FamilyID issues found"))

        # Determine FamilyID mapping to use
        familyid_mapping = None
        if familyid_policy == "auto-merge" and auto_merge_map:
            familyid_mapping = auto_merge_map
            self.stdout.write(
                self.style.WARNING(
                    f"  Auto-merge will remap {len(auto_merge_map)} FamilyIDs"
                )
            )
            self._write_familyid_merge_report(auto_merge_map, year, report_dir)

        # Import tblKartei
        self.stdout.write(self.style.NOTICE("\nImporting tblKartei..."))
        main_stats = import_tbl_kartei(
            tbl_kartei_rows,
            year,
            dry_run=dry_run,
            familyid_mapping=familyid_mapping,
        )
        self._print_stats("tblKartei", main_stats)

        # Import pre_tblKartei
        pending_stats = ImportStats()
        if not skip_pending and pre_rows:
            self.stdout.write(self.style.NOTICE("\nImporting pre_tblKartei (pending changes)..."))
            pending_stats = import_pre_tbl_kartei(
                pre_rows,
                year,
                dry_run=dry_run,
                familyid_mapping=familyid_mapping,
            )
            self._print_stats("pre_tblKartei", pending_stats)

        # Import decl_tblKartei
        declined_stats = ImportStats()
        if not skip_declined and decl_rows:
            self.stdout.write(self.style.NOTICE("\nImporting decl_tblKartei (declined changes)..."))
            declined_stats = import_decl_tbl_kartei(
                decl_rows,
                year,
                dry_run=dry_run,
                familyid_mapping=familyid_mapping,
            )
            self._print_stats("decl_tblKartei", declined_stats)

        # Sync history if requested
        history_events_created = 0
        if sync_history:
            self.stdout.write(self.style.NOTICE("\nSyncing history..."))
            history_events_created = sync_history_for_records(year, dry_run=dry_run)
            if dry_run:
                self.stdout.write(f"  Would create ~{history_events_created} history events")
            else:
                self.stdout.write(
                    self.style.SUCCESS(f"  Created {history_events_created} history events")
                )

        # Aggregate stats
        total_stats = ImportStats()
        total_stats.merge(main_stats)
        total_stats.merge(pending_stats)
        total_stats.merge(declined_stats)
        total_stats.familyid_issues = len(familyid_issues)

        # Write JSON report
        self._write_json_report(total_stats, year, report_dir, dry_run, history_events_created)

        # Final summary
        self.stdout.write(self.style.NOTICE(f"\n{'='*60}"))
        self.stdout.write(self.style.NOTICE("Import Summary"))
        self.stdout.write(self.style.NOTICE(f"{'='*60}"))
        self.stdout.write(f"Total rows processed: {total_stats.total_rows}")
        self.stdout.write(f"Records created: {total_stats.created_records}")
        self.stdout.write(f"Records updated: {total_stats.updated_records}")
        self.stdout.write(f"Records skipped: {total_stats.skipped_records}")
        self.stdout.write(f"Marker rows skipped: {total_stats.marker_skipped}")
        self.stdout.write(f"Pending changes created: {total_stats.pending_created}")
        self.stdout.write(f"Pending changes updated: {total_stats.pending_updated}")
        self.stdout.write(f"Declined changes created: {total_stats.declined_created}")
        self.stdout.write(f"Conflicts (year, id): {total_stats.conflicts_year_id}")
        self.stdout.write(f"FamilyID issues: {total_stats.familyid_issues}")
        self.stdout.write(f"Parse errors: {total_stats.parse_errors}")

        if dry_run:
            self.stdout.write(
                self.style.WARNING("\nDRY RUN completed. No changes were made.")
            )
        else:
            self.stdout.write(self.style.SUCCESS("\nImport completed successfully!"))

        # Report files created
        self.stdout.write(self.style.NOTICE("\nReports generated:"))
        self.stdout.write(f"  {report_dir / f'import_stats_{year}.json'}")
        if familyid_issues:
            self.stdout.write(f"  {report_dir / f'import_familyid_issues_{year}.csv'}")
        if familyid_mapping:
            self.stdout.write(f"  {report_dir / f'import_familyid_merge_{year}.csv'}")

    def _print_stats(self, table_name: str, stats: "ImportStats"):
        """Print import statistics for a table."""
        self.stdout.write(f"  Total rows: {stats.total_rows}")
        self.stdout.write(f"  Created: {stats.created_records}")
        self.stdout.write(f"  Updated: {stats.updated_records}")
        self.stdout.write(f"  Skipped: {stats.skipped_records}")
        self.stdout.write(f"  Marker rows skipped: {stats.marker_skipped}")
        if stats.conflicts_year_id:
            self.stdout.write(
                self.style.WARNING(f"  Conflicts (year, id): {stats.conflicts_year_id}")
            )
        if stats.pending_created or stats.pending_updated:
            self.stdout.write(
                f"  Pending: {stats.pending_created} created, {stats.pending_updated} updated"
            )
        if stats.declined_created:
            self.stdout.write(f"  Declined created: {stats.declined_created}")
        if stats.parse_errors:
            self.stdout.write(
                self.style.WARNING(f"  Parse errors: {stats.parse_errors}")
            )

    def _write_familyid_report(
        self,
        issues: list["FamilyIdIssue"],
        year: int,
        report_dir: Path,
    ):
        """Write FamilyID issues to CSV file."""
        filename = report_dir / f"import_familyid_issues_{year}.csv"

        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "issue_type",
                    "year",
                    "record_id",
                    "family_id",
                    "parent",
                    "email",
                    "family_key",
                    "details",
                ],
            )
            writer.writeheader()
            for issue in issues:
                writer.writerow({
                    "issue_type": issue.issue_type,
                    "year": issue.year,
                    "record_id": issue.record_id,
                    "family_id": issue.family_id,
                    "parent": issue.parent,
                    "email": issue.email,
                    "family_key": issue.family_key,
                    "details": issue.details,
                })

        self.stdout.write(f"  Wrote FamilyID issues report: {filename}")

    def _write_familyid_merge_report(
        self,
        merge_map: dict[str, str],
        year: int,
        report_dir: Path,
    ):
        """Write FamilyID merge mapping to CSV file."""
        filename = report_dir / f"import_familyid_merge_{year}.csv"

        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["old_family_id", "new_family_id"],
            )
            writer.writeheader()
            for old_fid, new_fid in sorted(merge_map.items()):
                writer.writerow({
                    "old_family_id": old_fid,
                    "new_family_id": new_fid,
                })

        self.stdout.write(f"  Wrote FamilyID merge report: {filename}")

    def _write_json_report(
        self,
        stats: "ImportStats",
        year: int,
        report_dir: Path,
        dry_run: bool,
        history_events: int,
    ):
        """Write full import report to JSON file."""
        filename = report_dir / f"import_stats_{year}.json"

        report = {
            "timestamp": datetime.now().isoformat(),
            "year": year,
            "dry_run": dry_run,
            "stats": stats.to_dict(),
            "history_events_created": history_events,
        }

        with open(filename, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2, ensure_ascii=False)

        self.stdout.write(f"  Wrote stats report: {filename}")

    # =========================================================================
    # Patch Mode: Update only specific fields without full re-import
    # =========================================================================

    def _handle_patch_mode(
        self,
        year: int,
        access_file: str,
        dry_run: bool,
        report_dir: Path,
    ):
        """
        Handle --patch-fields mode: update only teacher names, contract
        type/status, and sepa_marker for existing records.
        
        This mode:
        - Reads ONLY tblKartei (skips pre_tblKartei and decl_tblKartei)
        - Finds existing records by (year, id)
        - Updates only patch-specific fields without touching other data
        - Supports --dry-run for preview
        """
        report_dir.mkdir(parents=True, exist_ok=True)

        self.stdout.write(self.style.NOTICE(f"\n{'='*60}"))
        self.stdout.write(self.style.NOTICE("Access → Django PATCH Import"))
        self.stdout.write(self.style.NOTICE(f"{'='*60}"))
        self.stdout.write(f"Year: {year}")
        self.stdout.write(f"Access file: {access_file}")
        self.stdout.write(f"Dry run: {dry_run}")
        self.stdout.write(f"Report directory: {report_dir}")
        self.stdout.write("")
        self.stdout.write(self.style.WARNING(
            "PATCH MODE: Nur Lehrer, Vertragstyp, Status und SEPA-Marker werden aktualisiert.\n"
            "Andere Felder bleiben unverändert.\n"
        ))

        if dry_run:
            self.stdout.write(self.style.WARNING(
                "DRY RUN MODE - Keine Änderungen werden in die Datenbank geschrieben\n"
            ))

        # Import modules
        try:
            from apps.legacy_import.access_client import (
                resolve_access_file,
                open_access_connection,
                load_tbl_kartei,
            )
            from apps.legacy_import.services import (
                PatchStats,
                patch_tbl_kartei,
                PATCH_ONLY_FIELDS,
            )
        except ImportError as e:
            raise CommandError(f"Failed to import required modules: {e}")

        self._ensure_patch_schema_ready(set(PATCH_ONLY_FIELDS))

        # Resolve Access file path
        try:
            file_path = resolve_access_file(access_file)
        except FileNotFoundError as e:
            raise CommandError(str(e))
        except ValueError as e:
            raise CommandError(str(e))

        self.stdout.write(f"Resolved file path: {file_path}")

        # Open connection and load tblKartei only
        try:
            with open_access_connection(file_path) as conn:
                self.stdout.write(self.style.NOTICE("\nLade Daten aus Access (nur tblKartei)..."))
                tbl_kartei_rows = list(load_tbl_kartei(conn, year))
                self.stdout.write(f"  Geladen: {len(tbl_kartei_rows)} Zeilen aus tblKartei")

        except ImportError as e:
            raise CommandError(
                f"pyodbc is not installed. Install it with: pip install pyodbc\n"
                f"Error: {e}"
            )
        except Exception as e:
            raise CommandError(f"Failed to connect to Access database: {e}")

        # Run patch import
        self.stdout.write(self.style.NOTICE("\nPatch-Import läuft..."))
        patch_stats = patch_tbl_kartei(tbl_kartei_rows, year, dry_run=dry_run)

        # Print statistics
        self.stdout.write(self.style.NOTICE(f"\n{'='*60}"))
        self.stdout.write(self.style.NOTICE("Patch-Import Zusammenfassung"))
        self.stdout.write(self.style.NOTICE(f"{'='*60}"))
        self.stdout.write(f"Zeilen verarbeitet: {patch_stats.total_rows}")
        self.stdout.write(f"Datensätze gefunden: {patch_stats.records_found}")
        self.stdout.write(f"Datensätze aktualisiert: {patch_stats.records_updated}")
        self.stdout.write(f"Datensätze nicht gefunden: {patch_stats.records_not_found}")
        self.stdout.write(f"Marker-Zeilen übersprungen: {patch_stats.marker_skipped}")
        
        if patch_stats.parse_errors:
            self.stdout.write(
                self.style.WARNING(f"Fehler: {patch_stats.parse_errors}")
            )

        if patch_stats.not_found_ids:
            self.stdout.write(self.style.WARNING(
                f"  Nicht gefundene IDs (erste 20): {patch_stats.not_found_ids[:20]}"
            ))

        # Write JSON report
        filename = report_dir / f"patch_stats_{year}.json"
        report = {
            "timestamp": datetime.now().isoformat(),
            "year": year,
            "dry_run": dry_run,
            "mode": "patch-fields",
            "stats": patch_stats.to_dict(),
        }
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        self.stdout.write(f"\n  Report geschrieben: {filename}")

        if dry_run:
            self.stdout.write(
                self.style.WARNING("\nDRY RUN abgeschlossen. Keine Änderungen wurden vorgenommen.")
            )
        else:
            self.stdout.write(self.style.SUCCESS("\nPatch-Import erfolgreich abgeschlossen!"))
