"""Safe Python wrapper around the PowerShell import-workbook builder."""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

from .config import AppConfig


class WorkbookBuildError(RuntimeError):
    """Raised when a KiotViet import workbook cannot be built or verified."""


@dataclass(frozen=True, slots=True)
class BuildResult:
    output_path: str
    sha256: str
    file_size: int
    order_count: int
    visible_sheet_count: int
    mappings: list[dict[str, Any]]
    stdout: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": self.output_path,
            "sha256": self.sha256,
            "file_size": self.file_size,
            "order_count": self.order_count,
            "visible_sheet_count": self.visible_sheet_count,
            "mappings": self.mappings,
        }


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class WorkbookBuilder:
    """Invoke ``build_import_workbook.ps1`` with an auditable JSON payload."""

    def __init__(self, config: AppConfig | None = None) -> None:
        self.config = config or AppConfig()

    def build(
        self,
        orders: Sequence[Mapping[str, Any]],
        output_path: str | Path,
    ) -> BuildResult:
        normalized = self._validate_orders(orders)
        output = Path(output_path).expanduser().resolve()
        self._validate_runtime()
        output.parent.mkdir(parents=True, exist_ok=True)
        self.config.temp_dir.mkdir(parents=True, exist_ok=True)

        payload = {"orders": normalized}
        payload_path: Path | None = None
        # Excel SaveAs still has a much smaller practical path limit than
        # modern Windows.  Keep the same-volume temporary name intentionally
        # short; the random token is sufficient because each build is also
        # serialized by its destination batch.
        partial = output.with_name(f".ob_{uuid.uuid4().hex[:12]}.xlsx")
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                suffix=".orders.json",
                prefix="kiotviet_",
                dir=self.config.temp_dir,
                delete=False,
            ) as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)
                handle.flush()
                os.fsync(handle.fileno())
                payload_path = Path(handle.name)

            command = [
                str(self.config.powershell_path),
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(self.config.workbook_script_path),
                "-TemplatePath",
                str(self.config.template_path),
                "-OutputPath",
                str(partial),
                "-OrdersJsonPath",
                str(payload_path),
            ]
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            try:
                completed = subprocess.run(
                    command,
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=self.config.workbook_timeout_seconds,
                    creationflags=creationflags,
                )
            except subprocess.TimeoutExpired as exc:
                raise WorkbookBuildError(
                    "PowerShell workbook builder timed out; no import was attempted."
                ) from exc
            except OSError as exc:
                raise WorkbookBuildError(
                    f"Could not start PowerShell workbook builder: {exc}"
                ) from exc

            stdout = completed.stdout.strip()
            stderr = completed.stderr.strip()
            if completed.returncode != 0:
                detail = stderr or stdout or f"exit code {completed.returncode}"
                raise WorkbookBuildError(f"Workbook builder failed: {detail}")

            metadata = self._parse_builder_output(stdout)
            if metadata.get("success") is not True:
                raise WorkbookBuildError(
                    "Workbook builder did not report a successful result."
                )
            if not partial.is_file() or partial.stat().st_size <= 0:
                raise WorkbookBuildError(
                    "Workbook builder returned success but no output file was produced."
                )

            order_count = int(metadata.get("order_count", len(normalized)))
            sheet_count = int(
                metadata.get("visible_sheet_count", len(normalized))
            )
            mappings = [dict(item) for item in metadata.get("mappings", [])]
            if order_count != len(normalized) or sheet_count != len(normalized):
                raise WorkbookBuildError(
                    "Workbook verification failed: order/sheet count does not match the batch."
                )
            if len(mappings) != len(normalized):
                raise WorkbookBuildError(
                    "Workbook verification failed: sheet mappings are incomplete."
                )
            for expected_index, (order, mapping) in enumerate(
                zip(normalized, mappings, strict=True), start=1
            ):
                if int(mapping.get("sheet_index", 0)) != expected_index:
                    raise WorkbookBuildError("Workbook sheet mapping order is invalid.")
                if str(mapping.get("order_code", "")) != order["order_code"]:
                    raise WorkbookBuildError("Workbook order-code mapping is invalid.")

            # os.replace is atomic on the same volume and prevents readers from
            # observing a half-written XLSX.
            os.replace(partial, output)
            return BuildResult(
                output_path=str(output),
                sha256=sha256_file(output),
                file_size=output.stat().st_size,
                order_count=order_count,
                visible_sheet_count=sheet_count,
                mappings=mappings,
                stdout=stdout,
            )
        finally:
            if payload_path is not None:
                payload_path.unlink(missing_ok=True)
            partial.unlink(missing_ok=True)

    def _validate_runtime(self) -> None:
        if not self.config.powershell_path.is_file():
            raise WorkbookBuildError(
                f"PowerShell was not found: {self.config.powershell_path}"
            )
        if not self.config.workbook_script_path.is_file():
            raise WorkbookBuildError(
                f"Workbook builder script was not found: {self.config.workbook_script_path}"
            )
        if not self.config.template_path.is_file():
            raise WorkbookBuildError(
                f"KiotViet template was not found: {self.config.template_path}"
            )

    def _validate_orders(
        self, orders: Sequence[Mapping[str, Any]]
    ) -> list[dict[str, Any]]:
        if not 1 <= len(orders) <= self.config.max_orders_per_batch:
            raise WorkbookBuildError(
                f"A workbook must contain 1-{self.config.max_orders_per_batch} orders."
            )

        result: list[dict[str, Any]] = []
        seen_codes: set[str] = set()
        for order in orders:
            code = str(order.get("order_code", "")).strip()
            if not code:
                raise WorkbookBuildError("An order is missing order_code.")
            if code in seen_codes:
                raise WorkbookBuildError(f"Duplicate order in one batch: {code}")
            seen_codes.add(code)

            products_value = order.get("products", [])
            if not isinstance(products_value, Sequence) or isinstance(
                products_value, (str, bytes)
            ):
                raise WorkbookBuildError(f"Products for {code} must be a list.")
            products = [dict(product) for product in products_value]
            if not 1 <= len(products) <= self.config.max_products_per_order:
                raise WorkbookBuildError(
                    f"Order {code} must contain 1-{self.config.max_products_per_order} products."
                )
            for product in products:
                if not str(product.get("product_code", "")).strip():
                    raise WorkbookBuildError(f"Order {code} contains a product without a code.")

            result.append(
                {
                    "order_code": code,
                    "customer_code": str(order.get("customer_code", "")).strip(),
                    "customer_name": str(order.get("customer_name", "")).strip(),
                    "products": products,
                }
            )
        return result

    @staticmethod
    def _parse_builder_output(stdout: str) -> dict[str, Any]:
        if not stdout:
            raise WorkbookBuildError("Workbook builder returned no metadata.")
        candidates = [stdout, *reversed(stdout.splitlines())]
        for candidate in candidates:
            candidate = candidate.strip()
            if not candidate.startswith("{"):
                continue
            try:
                parsed = json.loads(candidate)
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict):
                return parsed
        raise WorkbookBuildError(
            "Workbook builder returned invalid JSON metadata."
        )


class PurchaseWorkbookBuilder:
    """Build the KiotViet purchase-import workbook from a shipment workbook.

    The source workbook is intentionally read by Excel COM because it may be
    an Office-generated workbook with formulas/formatting.  The PowerShell
    helper keeps the sample template formatting and filters only rows with a
    positive delivered quantity.  Unit prices from the source ``Giá nhập``
    column are divided by 1.08 before upload;
    VAT is enabled in KiotViet after the import.
    """

    def __init__(self, config: AppConfig | None = None) -> None:
        self.config = config or AppConfig()

    def build(
        self,
        source_path: str | Path,
        output_path: str | Path,
        *,
        baseline_path: str | Path | None = None,
    ) -> dict[str, Any]:
        source = Path(source_path).expanduser().resolve()
        output = Path(output_path).expanduser().resolve()
        baseline = (
            Path(baseline_path).expanduser().resolve()
            if baseline_path is not None and str(baseline_path).strip()
            else None
        )
        template = Path(self.config.purchase_template_path)
        script = self.config.scripts_dir / "build_purchase_workbook.ps1"
        if not source.is_file() or source.name.startswith("~$"):
            raise WorkbookBuildError(f"Không tìm thấy file nguồn: {source}")
        if baseline is not None:
            if not baseline.is_file() or baseline.name.startswith("~$"):
                raise WorkbookBuildError(f"Không tìm thấy file đối chiếu: {baseline}")
            if baseline == source:
                raise WorkbookBuildError("File đối chiếu phải khác file nguồn cần nhập.")
        if not template.is_file():
            raise WorkbookBuildError(f"Không tìm thấy file mẫu nhập hàng: {template}")
        if not script.is_file():
            raise WorkbookBuildError(f"Không tìm thấy script tạo file nhập hàng: {script}")
        output.parent.mkdir(parents=True, exist_ok=True)
        command = [
            str(self.config.powershell_path), "-NoProfile", "-NonInteractive",
            "-ExecutionPolicy", "Bypass", "-File", str(script),
            "-TemplatePath", str(template), "-SourcePath", str(source),
            "-OutputPath", str(output), "-VatDivisor", "1.08",
        ]
        if baseline is not None:
            command.extend(["-BaselinePath", str(baseline)])
        try:
            completed = subprocess.run(
                command, capture_output=True, text=True, encoding="utf-8",
                errors="replace", timeout=self.config.workbook_timeout_seconds,
                check=False, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as exc:
            raise WorkbookBuildError("Tạo file nhập hàng quá thời gian; chưa upload.") from exc
        if completed.returncode != 0:
            raise WorkbookBuildError(completed.stderr.strip() or completed.stdout.strip() or "Không tạo được file nhập hàng.")
        metadata = WorkbookBuilder._parse_builder_output(completed.stdout.strip())
        if metadata.get("success") is not True or not output.is_file():
            raise WorkbookBuildError("Script không xác nhận file nhập hàng hợp lệ.")
        metadata["output_path"] = str(output)
        if baseline is not None:
            metadata["baseline_path"] = str(baseline)
        metadata["sha256"] = sha256_file(output)
        metadata["file_size"] = output.stat().st_size
        return metadata


__all__ = [
    "BuildResult",
    "WorkbookBuildError",
    "WorkbookBuilder",
    "PurchaseWorkbookBuilder",
    "sha256_file",
]
