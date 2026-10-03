"""Application configuration for the local Sopoka -> KiotViet workflow."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path


# Thailand has used UTC+07:00 without daylight-saving changes for this
# workflow.  A fixed offset avoids requiring the optional ``tzdata`` wheel on
# Windows while still making the business-date boundary explicit.
BANGKOK_TZ = timezone(timedelta(hours=7), name="Asia/Bangkok")


def _env_path(name: str, fallback: Path) -> Path:
    raw = os.environ.get(name)
    return Path(raw).expanduser().resolve() if raw else fallback.resolve()


@dataclass(frozen=True, slots=True)
class AppConfig:
    """Resolved paths and immutable workflow limits.

    All defaults stay inside ``Desktop/Automation``.  Environment overrides
    are primarily useful for tests and portable deployments.
    """

    app_root: Path = field(
        default_factory=lambda: Path(__file__).resolve().parents[1]
    )
    data_dir: Path | None = None
    outputs_dir: Path | None = None
    logs_dir: Path | None = None
    resources_dir: Path | None = None
    scripts_dir: Path | None = None
    temp_dir: Path | None = None
    database_path: Path | None = None
    template_path: Path | None = None
    workbook_script_path: Path | None = None
    powershell_path: Path | None = None
    edge_driver_path: Path | None = None
    edge_binary_path: Path | None = None
    edge_profile_dir: Path | None = None
    edge_download_dir: Path | None = None
    purchase_template_path: Path | None = None

    expected_branch: str = "Đồng Nai 2 - ĐN2"
    customer_code_suffix: str = "DN1"
    max_orders_per_batch: int = 19
    max_products_per_order: int = 250
    sqlite_busy_timeout_ms: int = 15_000
    workbook_timeout_seconds: int = 180
    scheduler_interval_hours: int = 2
    scheduler_start_hour: int = 7
    scheduler_end_hour: int = 21
    scheduler_start_minute: int = 5
    scheduler_max_clock_skew_seconds: int = 120

    def __post_init__(self) -> None:
        root = Path(self.app_root).expanduser().resolve()
        object.__setattr__(self, "app_root", root)

        data = Path(self.data_dir).resolve() if self.data_dir else _env_path(
            "AUTOMATION_DATA_DIR", root / "data"
        )
        outputs = (
            Path(self.outputs_dir).resolve()
            if self.outputs_dir
            else _env_path("AUTOMATION_OUTPUTS_DIR", root / "outputs")
        )
        logs = Path(self.logs_dir).resolve() if self.logs_dir else _env_path(
            "AUTOMATION_LOGS_DIR", root / "logs"
        )
        resources = (
            Path(self.resources_dir).resolve()
            if self.resources_dir
            else _env_path("AUTOMATION_RESOURCES_DIR", root / "resources")
        )
        scripts = (
            Path(self.scripts_dir).resolve()
            if self.scripts_dir
            else _env_path("AUTOMATION_SCRIPTS_DIR", root / "scripts")
        )
        temporary = (
            Path(self.temp_dir).resolve()
            if self.temp_dir
            else _env_path("AUTOMATION_TEMP_DIR", data / "tmp")
        )

        object.__setattr__(self, "data_dir", data)
        object.__setattr__(self, "outputs_dir", outputs)
        object.__setattr__(self, "logs_dir", logs)
        object.__setattr__(self, "resources_dir", resources)
        object.__setattr__(self, "scripts_dir", scripts)
        object.__setattr__(self, "temp_dir", temporary)

        database = (
            Path(self.database_path).resolve()
            if self.database_path
            else _env_path("AUTOMATION_DB_PATH", data / "automation.sqlite3")
        )
        template = (
            Path(self.template_path).resolve()
            if self.template_path
            else _env_path(
                "AUTOMATION_TEMPLATE_PATH",
                resources / "KiotVietImportTemplate.xlsx",
            )
        )
        workbook_script = (
            Path(self.workbook_script_path).resolve()
            if self.workbook_script_path
            else _env_path(
                "AUTOMATION_WORKBOOK_SCRIPT",
                scripts / "build_import_workbook.ps1",
            )
        )
        powershell = (
            Path(self.powershell_path).resolve()
            if self.powershell_path
            else _env_path(
                "AUTOMATION_POWERSHELL",
                Path(os.environ.get("SystemRoot", r"C:\Windows"))
                / "System32"
                / "WindowsPowerShell"
                / "v1.0"
                / "powershell.exe",
            )
        )
        object.__setattr__(self, "database_path", database)
        object.__setattr__(self, "template_path", template)
        object.__setattr__(self, "workbook_script_path", workbook_script)
        object.__setattr__(self, "powershell_path", powershell)

        edge_driver = (
            Path(self.edge_driver_path).resolve()
            if self.edge_driver_path
            else _env_path("AUTOMATION_EDGE_DRIVER", root / "msedgedriver.exe")
        )
        edge_binary = (
            Path(self.edge_binary_path).resolve()
            if self.edge_binary_path
            else _env_path(
                "AUTOMATION_EDGE_BINARY",
                Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
                / "Microsoft"
                / "Edge"
                / "Application"
                / "msedge.exe",
            )
        )
        edge_profile = (
            Path(self.edge_profile_dir).resolve()
            if self.edge_profile_dir
            else _env_path("AUTOMATION_EDGE_PROFILE", data / "edge-profile")
        )
        edge_downloads = (
            Path(self.edge_download_dir).resolve()
            if self.edge_download_dir
            else _env_path("AUTOMATION_EDGE_DOWNLOADS", data / "downloads")
        )
        object.__setattr__(self, "edge_driver_path", edge_driver)
        object.__setattr__(self, "edge_binary_path", edge_binary)
        object.__setattr__(self, "edge_profile_dir", edge_profile)
        object.__setattr__(self, "edge_download_dir", edge_downloads)
        purchase_template = (
            Path(self.purchase_template_path).resolve()
            if self.purchase_template_path
            else _env_path(
                "AUTOMATION_PURCHASE_TEMPLATE",
                Path(r"C:\Users\binni\Downloads\MauFileNhapHang (4).xlsx"),
            )
        )
        object.__setattr__(self, "purchase_template_path", purchase_template)

        if not 1 <= self.max_orders_per_batch <= 19:
            raise ValueError("max_orders_per_batch must be between 1 and 19")
        if not 1 <= self.max_products_per_order <= 250:
            raise ValueError("max_products_per_order must be between 1 and 250")
        if not 1 <= self.scheduler_interval_hours <= 24:
            raise ValueError("scheduler_interval_hours must be between 1 and 24")
        if not 0 <= self.scheduler_start_hour <= 23:
            raise ValueError("scheduler_start_hour must be between 0 and 23")
        if not 0 <= self.scheduler_end_hour <= 23:
            raise ValueError("scheduler_end_hour must be between 0 and 23")
        if self.scheduler_end_hour < self.scheduler_start_hour:
            raise ValueError("scheduler_end_hour must not be before scheduler_start_hour")
        if not 0 <= self.scheduler_start_minute <= 59:
            raise ValueError("scheduler_start_minute must be between 0 and 59")
        if not 1 <= self.scheduler_max_clock_skew_seconds <= 900:
            raise ValueError("scheduler_max_clock_skew_seconds must be between 1 and 900")

    def ensure_directories(self) -> None:
        """Create runtime-owned directories, never source/template files."""

        for path in (
            self.data_dir,
            self.outputs_dir,
            self.logs_dir,
            self.temp_dir,
            self.edge_profile_dir,
            self.edge_download_dir,
            self.database_path.parent,
        ):
            path.mkdir(parents=True, exist_ok=True)

    def now(self) -> datetime:
        """Return the current application time in the business timezone."""

        return datetime.now(BANGKOK_TZ)

    def today(self) -> date:
        """Return today's Thailand business date."""

        return self.now().date()


__all__ = ["AppConfig", "BANGKOK_TZ"]
