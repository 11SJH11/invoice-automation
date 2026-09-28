#!/usr/bin/env python3
"""
Fingertips Invoice Automation

This is the first integrated desktop interface for the existing automation
pipeline. It deliberately ORCHESTRATES the proven scripts rather than
duplicating/re-writing all of their logic.

Expected files beside this script
---------------------------------
email_cache.py
mydocsonline_import.py
email_classifier.py
spreadsheet_generator.py
invoice_generator.py             (optional until invoice stage)
logo.png                           (optional; invoice generator can fall back)

Main workflow
-------------
Choose month
    -> Fetch/cache new email
    -> Fetch/parse MyDocsOnline notifications
    -> Classify cached email
    -> Generate monthly master spreadsheet
    -> User reviews spreadsheet
    -> User clicks Generate Invoices when ready

Credentials
-----------
The UI can take the mailbox username/password for the current run only.
They are passed to child Python processes as environment variables and are
NOT written to disk by this interface.

You may also set:
    FINGERTIPS_EMAIL
    FINGERTIPS_EMAIL_PASSWORD

Run:
    python fingertips_app.py
"""

from __future__ import annotations

import calendar
import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Fingertips Invoice Automation"

SCRIPT_DIR = Path(__file__).resolve().parent

CACHE_SCRIPT = SCRIPT_DIR / "email_cache.py"
MDO_SCRIPT = SCRIPT_DIR / "mydocsonline_import.py"
CLASSIFIER_SCRIPT = SCRIPT_DIR / "email_classifier.py"
SPREADSHEET_SCRIPT = SCRIPT_DIR / "spreadsheet_generator.py"
INVOICE_SCRIPT = SCRIPT_DIR / "invoice_generator.py"

OUTPUT_DIR = SCRIPT_DIR / "generated_spreadsheets"
INVOICE_DIR = SCRIPT_DIR / "generated_invoices"
RATE_OVERRIDES = SCRIPT_DIR / "rate_overrides.csv"

REQUIRED_PIPELINE_SCRIPTS = [
    CACHE_SCRIPT,
    MDO_SCRIPT,
    CLASSIFIER_SCRIPT,
    SPREADSHEET_SCRIPT,
]


def month_to_label(month_value: str) -> str:
    dt = datetime.strptime(month_value, "%Y-%m")
    return f"{calendar.month_name[dt.month]}_{dt.year}"


def expected_workbook_path(month_value: str) -> Path:
    return OUTPUT_DIR / f"Fingertips_Jobs_{month_to_label(month_value)}.xlsx"


def open_path(path: Path) -> None:
    path = path.resolve()

    if not path.exists():
        raise FileNotFoundError(path)

    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class FingertipsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("980x760")
        self.minsize(900, 680)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.last_workbook: Path | None = None

        self.month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        self.email_var = tk.StringVar(
            value=os.getenv(
                "FINGERTIPS_EMAIL",
                "cbennett@fingertipstyping.co.uk",
            )
        )
        self.password_var = tk.StringVar(
            value=os.getenv("FINGERTIPS_EMAIL_PASSWORD", "")
        )

        self.status_var = tk.StringVar(value="Ready")
        self.workbook_var = tk.StringVar(value="No spreadsheet generated this session.")

        self._build_ui()
        self._refresh_preflight()
        self.after(100, self._drain_log_queue)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(
            outer,
            text="Fingertips Invoice Automation",
            font=("Segoe UI", 20, "bold"),
        )
        title.pack(anchor="w")

        subtitle = ttk.Label(
            outer,
            text=(
                "Fetch work → analyse emails → generate monthly spreadsheet "
                "→ review → generate invoices"
            ),
        )
        subtitle.pack(anchor="w", pady=(2, 16))

        settings = ttk.LabelFrame(outer, text="Monthly run", padding=12)
        settings.pack(fill="x")

        ttk.Label(settings, text="Invoice month (YYYY-MM):").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=5
        )

        month_entry = ttk.Entry(
            settings,
            textvariable=self.month_var,
            width=14,
        )
        month_entry.grid(row=0, column=1, sticky="w", pady=5)

        ttk.Label(settings, text="Mailbox email:").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=5
        )
        ttk.Entry(
            settings,
            textvariable=self.email_var,
            width=42,
        ).grid(row=1, column=1, sticky="w", pady=5)

        ttk.Label(settings, text="Mailbox password:").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=5
        )
        password_entry = ttk.Entry(
            settings,
            textvariable=self.password_var,
            show="•",
            width=42,
        )
        password_entry.grid(row=2, column=1, sticky="w", pady=5)

        ttk.Label(
            settings,
            text="Password is used for this run only and is not saved by the interface.",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(0, 4))

        actions = ttk.LabelFrame(outer, text="Actions", padding=12)
        actions.pack(fill="x", pady=(14, 0))

        self.full_run_btn = ttk.Button(
            actions,
            text="Run Monthly Import",
            command=self.run_full_pipeline,
        )
        self.full_run_btn.grid(row=0, column=0, padx=(0, 8), pady=4, sticky="ew")

        self.cache_run_btn = ttk.Button(
            actions,
            text="Build From Existing Cache",
            command=self.run_cached_pipeline,
        )
        self.cache_run_btn.grid(row=0, column=1, padx=8, pady=4, sticky="ew")

        self.open_sheet_btn = ttk.Button(
            actions,
            text="Open Spreadsheet",
            command=self.open_spreadsheet,
        )
        self.open_sheet_btn.grid(row=0, column=2, padx=8, pady=4, sticky="ew")

        self.invoice_btn = ttk.Button(
            actions,
            text="Generate Invoices",
            command=self.generate_invoices,
        )
        self.invoice_btn.grid(row=1, column=0, padx=(0, 8), pady=4, sticky="ew")

        ttk.Button(
            actions,
            text="Open Invoice Folder",
            command=self.open_invoice_folder,
        ).grid(row=1, column=1, padx=8, pady=4, sticky="ew")

        ttk.Button(
            actions,
            text="Open Rate Overrides",
            command=self.open_rate_overrides,
        ).grid(row=1, column=2, padx=8, pady=4, sticky="ew")

        for c in range(3):
            actions.columnconfigure(c, weight=1)

        output_frame = ttk.LabelFrame(
            outer,
            text="Current output",
            padding=12,
        )
        output_frame.pack(fill="x", pady=(14, 0))

        ttk.Label(
            output_frame,
            textvariable=self.workbook_var,
            wraplength=880,
        ).pack(anchor="w")

        preflight = ttk.LabelFrame(
            outer,
            text="System check",
            padding=12,
        )
        preflight.pack(fill="x", pady=(14, 0))

        self.preflight_text = tk.Text(
            preflight,
            height=6,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
        )
        self.preflight_text.pack(fill="x")

        log_frame = ttk.LabelFrame(
            outer,
            text="Run log",
            padding=8,
        )
        log_frame.pack(fill="both", expand=True, pady=(14, 0))

        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
        )
        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        status = ttk.Label(
            outer,
            textvariable=self.status_var,
            anchor="w",
        )
        status.pack(fill="x", pady=(8, 0))

    # ------------------------------------------------------------------
    # General helpers
    # ------------------------------------------------------------------

    def _set_text(self, widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", text)
        widget.configure(state="disabled")

    def _log(self, message: str) -> None:
        self.log_queue.put(message.rstrip() + "\n")

    def _drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break

            self.log_text.configure(state="normal")
            self.log_text.insert("end", message)
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

        self.after(100, self._drain_log_queue)

    def _validate_month(self) -> str | None:
        value = self.month_var.get().strip()
        try:
            datetime.strptime(value, "%Y-%m")
        except ValueError:
            messagebox.showerror(
                "Invalid month",
                "Enter the month as YYYY-MM, for example 2026-08.",
            )
            return None
        return value

    def _set_busy(self, busy: bool, status: str = "") -> None:
        state = "disabled" if busy else "normal"

        for widget in (
            self.full_run_btn,
            self.cache_run_btn,
            self.invoice_btn,
        ):
            widget.configure(state=state)

        self.status_var.set(
            status if status else ("Working..." if busy else "Ready")
        )

    def _run_background(self, target, status: str) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo(
                "Already running",
                "A Fingertips automation task is already running.",
            )
            return

        self._set_busy(True, status)

        def wrapper() -> None:
            try:
                target()
            except Exception:
                self._log("\nERROR")
                self._log(traceback.format_exc())
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Automation failed",
                        "The run failed. Check the Run log for details.",
                    ),
                )
                self.after(0, lambda: self._set_busy(False, "Run failed"))
            else:
                self.after(0, lambda: self._set_busy(False, "Ready"))
            finally:
                self.after(0, self._refresh_preflight)

        self.worker = threading.Thread(target=wrapper, daemon=True)
        self.worker.start()

    def _refresh_preflight(self) -> None:
        lines = []

        for path in REQUIRED_PIPELINE_SCRIPTS:
            marker = "OK" if path.exists() else "MISSING"
            lines.append(f"{marker:7} {path.name}")

        invoice_marker = "OK" if INVOICE_SCRIPT.exists() else "OPTIONAL/MISSING"
        lines.append(f"{invoice_marker:16} {INVOICE_SCRIPT.name}")

        try:
            import openpyxl  # noqa: F401
            lines.append("OK      openpyxl")
        except Exception:
            lines.append("MISSING openpyxl (run: python -m pip install openpyxl)")

        if self.password_var.get() or os.getenv("FINGERTIPS_EMAIL_PASSWORD"):
            lines.append("OK      mailbox credentials available for this session")
        else:
            lines.append("INFO    mailbox password not entered (cache-only mode still works)")

        workbook = expected_workbook_path(self.month_var.get().strip()) if self._month_looks_valid() else None
        if workbook and workbook.exists():
            self.last_workbook = workbook
            self.workbook_var.set(str(workbook.resolve()))

        self._set_text(self.preflight_text, "\n".join(lines))

    def _month_looks_valid(self) -> bool:
        try:
            datetime.strptime(self.month_var.get().strip(), "%Y-%m")
            return True
        except ValueError:
            return False

    def _check_required_scripts(self, include_fetch: bool = True) -> bool:
        required = [MDO_SCRIPT, CLASSIFIER_SCRIPT, SPREADSHEET_SCRIPT]
        if include_fetch:
            required.insert(0, CACHE_SCRIPT)

        missing = [path.name for path in required if not path.exists()]

        if missing:
            messagebox.showerror(
                "Missing scripts",
                "Place these scripts in the same folder as the interface:\n\n"
                + "\n".join(missing),
            )
            return False

        return True

    def _child_env(self) -> dict[str, str]:
        env = os.environ.copy()

        email_value = self.email_var.get().strip()
        password_value = self.password_var.get()

        if email_value:
            env["FINGERTIPS_EMAIL"] = email_value

        if password_value:
            env["FINGERTIPS_EMAIL_PASSWORD"] = password_value

        # Keep child-process logs Unicode-safe on Windows.
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        return env

    def _run_command(
        self,
        title: str,
        args: list[str],
        env: dict[str, str] | None = None,
    ) -> None:
        self._log("")
        self._log("=" * 78)
        self._log(title)
        self._log("=" * 78)
        self._log("Running: " + " ".join(args))

        process = subprocess.Popen(
            args,
            cwd=str(SCRIPT_DIR),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )

        assert process.stdout is not None

        for line in process.stdout:
            self._log(line.rstrip())

        return_code = process.wait()

        if return_code != 0:
            raise RuntimeError(
                f"{title} failed with exit code {return_code}."
            )

        self._log(f"{title}: COMPLETE")

    def _run_command_capture(
        self,
        title: str,
        args: list[str],
        env: dict[str, str] | None = None,
    ) -> tuple[int, str]:
        """
        Run a short child-process check and capture its output.

        Used for invoice warning preflight so the GUI can ask the user whether
        to continue when the data is suspicious but not actually invalid.
        """
        self._log("")
        self._log("=" * 78)
        self._log(title)
        self._log("=" * 78)
        self._log("Running: " + " ".join(args))

        completed = subprocess.run(
            args,
            cwd=str(SCRIPT_DIR),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        output = completed.stdout or ""

        for line in output.splitlines():
            self._log(line)

        return completed.returncode, output

    # ------------------------------------------------------------------
    # Pipeline
    # ------------------------------------------------------------------

    def run_full_pipeline(self) -> None:
        month = self._validate_month()
        if not month:
            return

        if not self._check_required_scripts(include_fetch=True):
            return

        if not self.email_var.get().strip() or not (
            self.password_var.get()
            or os.getenv("FINGERTIPS_EMAIL_PASSWORD")
        ):
            messagebox.showerror(
                "Mailbox credentials required",
                "Enter the mailbox email and password before fetching new email.\n\n"
                "Use 'Build From Existing Cache' if you only want to test the "
                "data already downloaded.",
            )
            return

        self._run_background(
            lambda: self._pipeline(month, fetch=True),
            f"Running monthly import for {month}...",
        )

    def run_cached_pipeline(self) -> None:
        month = self._validate_month()
        if not month:
            return

        if not self._check_required_scripts(include_fetch=False):
            return

        self._run_background(
            lambda: self._pipeline(month, fetch=False),
            f"Building {month} from existing cache...",
        )

    def _pipeline(self, month: str, fetch: bool) -> None:
        env = self._child_env()
        workbook = expected_workbook_path(month)
        workbook.parent.mkdir(parents=True, exist_ok=True)

        self._log("")
        self._log("#" * 78)
        self._log(f"FINGERTIPS MONTHLY PIPELINE - {month}")
        self._log("#" * 78)

        if fetch:
            self._run_command(
                "1/4 Cache new mailbox email",
                [sys.executable, str(CACHE_SCRIPT)],
                env,
            )

            self._run_command(
                "2/4 Fetch and parse MyDocsOnline",
                [
                    sys.executable,
                    str(MDO_SCRIPT),
                    "--month",
                    month,
                    "--output",
                    str(
                        SCRIPT_DIR
                        / "email_cache"
                        / "mydocsonline"
                        / f"MDO_{month}_intermediate.xlsx"
                    ),
                ],
                env,
            )
        else:
            self._log("")
            self._log("Using existing local mailbox cache (no server fetch).")

            self._run_command(
                "1/3 Re-parse cached MyDocsOnline email",
                [
                    sys.executable,
                    str(MDO_SCRIPT),
                    "--parse-only",
                    "--month",
                    month,
                    "--output",
                    str(
                        SCRIPT_DIR
                        / "email_cache"
                        / "mydocsonline"
                        / f"MDO_{month}_intermediate.xlsx"
                    ),
                ],
                env,
            )

        self._run_command(
            "Classify cached email",
            [sys.executable, str(CLASSIFIER_SCRIPT)],
            env,
        )

        self._run_command(
            "Generate monthly master spreadsheet",
            [
                sys.executable,
                str(SPREADSHEET_SCRIPT),
                "--month",
                month,
                "--output",
                str(workbook),
            ],
            env,
        )

        if not workbook.exists():
            raise RuntimeError(
                f"Pipeline completed but expected workbook was not found: {workbook}"
            )

        self.last_workbook = workbook

        self.after(
            0,
            lambda p=workbook: self.workbook_var.set(str(p.resolve())),
        )

        self._log("")
        self._log("#" * 78)
        self._log("MONTHLY PIPELINE COMPLETE")
        self._log(f"Spreadsheet: {workbook.resolve()}")
        self._log(
            "NEXT: Open the spreadsheet, review Needs Review / inferred rates, "
            "save your corrections, then return here and click Generate Invoices."
        )
        self._log("#" * 78)

        self.after(
            0,
            lambda: messagebox.showinfo(
                "Spreadsheet ready",
                "The monthly spreadsheet has been generated.\n\n"
                "Review and save it before generating invoices.",
            ),
        )

    # ------------------------------------------------------------------
    # Output / invoice actions
    # ------------------------------------------------------------------

    def _resolve_current_workbook(self) -> Path | None:
        month = self._validate_month()
        if not month:
            return None

        if self.last_workbook and self.last_workbook.exists():
            return self.last_workbook

        expected = expected_workbook_path(month)
        if expected.exists():
            self.last_workbook = expected
            return expected

        return None

    def open_spreadsheet(self) -> None:
        workbook = self._resolve_current_workbook()

        if not workbook:
            messagebox.showinfo(
                "No spreadsheet",
                "Generate the spreadsheet first.",
            )
            return

        try:
            open_path(workbook)
        except Exception as exc:
            messagebox.showerror("Could not open spreadsheet", str(exc))

    def generate_invoices(self) -> None:
        workbook = self._resolve_current_workbook()

        if not workbook:
            messagebox.showerror(
                "No spreadsheet",
                "Generate and review the monthly spreadsheet first.",
            )
            return

        if not INVOICE_SCRIPT.exists():
            messagebox.showerror(
                "Invoice generator missing",
                f"Cannot find {INVOICE_SCRIPT.name} beside the interface.",
            )
            return

        confirmed = messagebox.askyesno(
            "Generate invoices?",
            "Have you reviewed and SAVED the spreadsheet?\n\n"
            "For best results, close the spreadsheet in Excel before continuing.\n\n"
            "Missing required data will block the affected client. "
            "Suspicious values will be shown as warnings and can be overridden.",
        )

        if not confirmed:
            return

        # Warning preflight is deliberately quick and runs before the longer
        # invoice-generation worker. Exit code:
        #   0 = no warnings
        #   2 = suspicious data found, user may override
        #   1 = actual check failure
        return_code, output = self._run_command_capture(
            "Invoice warning preflight",
            [
                sys.executable,
                str(INVOICE_SCRIPT),
                str(workbook),
                "--output",
                str(INVOICE_DIR),
                "--check-warnings",
            ],
            self._child_env(),
        )

        allow_warnings = False

        if return_code == 2:
            warning_lines = [
                line for line in output.splitlines()
                if line.startswith("[WARNING]") or line.strip().startswith("- ")
            ]

            preview = "\n".join(warning_lines[:18])

            if len(warning_lines) > 18:
                preview += f"\n... and {len(warning_lines) - 18} more warning line(s)."

            proceed = messagebox.askyesno(
                "Suspicious invoice data found",
                "Some values look unusual but are not necessarily wrong.\n\n"
                f"{preview}\n\n"
                "Generate invoices anyway?",
            )

            if not proceed:
                self.status_var.set("Invoice generation cancelled - warnings require review")
                return

            allow_warnings = True

        elif return_code != 0:
            messagebox.showerror(
                "Invoice preflight failed",
                "The invoice-data check failed. Review the Run log for details.",
            )
            return

        self._run_background(
            lambda: self._invoice_pipeline(workbook, allow_warnings),
            "Generating invoices...",
        )

    def _invoice_pipeline(self, workbook: Path, allow_warnings: bool = False) -> None:
        command = [
            sys.executable,
            str(INVOICE_SCRIPT),
            str(workbook),
            "--output",
            str(INVOICE_DIR),
        ]

        if allow_warnings:
            command.append("--allow-warnings")

        self._run_command(
            "Generate invoices and schedules",
            command,
            self._child_env(),
        )

        self._log("")
        self._log(f"Invoice output: {INVOICE_DIR.resolve()}")

        self.after(
            0,
            lambda: messagebox.showinfo(
                "Invoice generation complete",
                "Invoice generation has finished.\n\n"
                "Check the Run log and generated_invoices folder before sending anything.",
            ),
        )

    def open_invoice_folder(self) -> None:
        INVOICE_DIR.mkdir(parents=True, exist_ok=True)
        try:
            open_path(INVOICE_DIR)
        except Exception as exc:
            messagebox.showerror("Could not open folder", str(exc))

    def open_rate_overrides(self) -> None:
        if not RATE_OVERRIDES.exists():
            RATE_OVERRIDES.write_text(
                "Client,Client Rate,Typist Rate,Notes,Active\n",
                encoding="utf-8-sig",
            )

        try:
            open_path(RATE_OVERRIDES)
        except Exception as exc:
            messagebox.showerror("Could not open rate overrides", str(exc))


def main() -> None:
    app = FingertipsApp()
    app.mainloop()


if __name__ == "__main__":
    main()
