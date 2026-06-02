"""Build the downloadable Excel workbook for reconciliation results.

Keeps all openpyxl styling concerns isolated from the Streamlit entry point.
"""
from __future__ import annotations

import io

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def _excel_money(value):
    """Coerce a value to a float amount, leaving non-numeric values untouched."""
    if value is None or value == "":
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        if isinstance(value, str):
            cleaned = value.replace("$", "").replace(",", "").strip()
            return float(cleaned)
        return float(value)
    except Exception:
        return value


def _expand_list_columns(df: pd.DataFrame, column_map: dict[str, str]) -> pd.DataFrame:
    """Explode list-valued columns into ``<prefix> 1``, ``<prefix> 2`` … columns."""
    if df is None or df.empty:
        return df

    df_out = df.copy()
    for source_col, prefix in column_map.items():
        if source_col not in df_out.columns:
            continue

        max_len = 0
        normalized_values = []
        for value in df_out[source_col].tolist():
            if isinstance(value, (list, tuple, set)):
                items = list(value)
            elif value is None or (isinstance(value, float) and pd.isna(value)):
                items = []
            else:
                items = [value]
            normalized_values.append(items)
            max_len = max(max_len, len(items))

        for idx in range(max_len):
            new_col = f"{prefix} {idx + 1}"
            df_out[new_col] = [
                items[idx] if idx < len(items) else None for items in normalized_values
            ]

        df_out = df_out.drop(columns=[source_col])

    return df_out


def _apply_currency_format(worksheet, header_prefixes: tuple[str, ...]) -> None:
    """Format numeric cells under matching headers as Colombian-peso currency."""
    header_map = {}
    for cell in worksheet[1]:
        if cell.value is not None:
            header_map[str(cell.value)] = cell.column

    for header, column_index in header_map.items():
        if not header.startswith(header_prefixes):
            continue

        column_letter = get_column_letter(column_index)
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet[f"{column_letter}{row_idx}"]
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right")


def _apply_row_status_styles(worksheet) -> None:
    """Colour each row green/red based on its ``Estado`` column."""
    header_map = {}
    for cell in worksheet[1]:
        if cell.value is not None:
            header_map[str(cell.value)] = cell.column

    estado_col = header_map.get("Estado")
    if estado_col is None:
        return

    ok_fill = PatternFill(fill_type="solid", fgColor="D4EDDA")
    ok_font = Font(color="155724")
    error_fill = PatternFill(fill_type="solid", fgColor="F8D7DA")
    error_font = Font(color="721C24")

    for row_idx in range(2, worksheet.max_row + 1):
        cell = worksheet.cell(row=row_idx, column=estado_col)
        estado = str(cell.value or "").strip().lower()
        if estado == "ok":
            fill, font = ok_fill, ok_font
        else:
            fill, font = error_fill, error_font

        for col_idx in range(1, worksheet.max_column + 1):
            current = worksheet.cell(row=row_idx, column=col_idx)
            current.fill = fill
            current.font = font


def build_reconciliation_excel_bytes(
    df_transfers: pd.DataFrame | None,
    df_seguridad: pd.DataFrame | None,
    recon_mode: str,
) -> bytes:
    """Return the styled reconciliation workbook as raw bytes for download."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if recon_mode == "transfers":
            sheet_name = "Transferencias"
            df_export = _expand_list_columns(
                df_transfers,
                {
                    "Neto_desprendibles": "Neto",
                    "Valores_transferencia": "Transferencia",
                },
            )
            if df_transfers is None or df_transfers.empty:
                pd.DataFrame(
                    columns=["Identificación", "Cuenta", "Estado", "Neto 1", "Transferencia 1"]
                ).to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df_export.to_excel(writer, sheet_name=sheet_name, index=False)
        elif recon_mode == "seguridad":
            sheet_name = "Seguridad_Social"
            df_export = df_seguridad.copy() if df_seguridad is not None else None
            if df_seguridad is None or df_seguridad.empty:
                pd.DataFrame(
                    columns=["Identificación", "Estado", "Devengado", "IBC"]
                ).to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                if df_export is not None:
                    df_export["Devengado"] = df_export["Devengado"].apply(_excel_money)
                    df_export["IBC"] = df_export["IBC"].apply(
                        lambda value: ", ".join(str(item) for item in value)
                        if isinstance(value, (list, tuple, set))
                        else _excel_money(value)
                    )
                df_export.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            if df_transfers is not None:
                df_transfers.to_excel(writer, sheet_name="Transferencias", index=False)
            if df_seguridad is not None:
                df_seguridad.to_excel(writer, sheet_name="Seguridad_Social", index=False)

        workbook = writer.book
        if "Transferencias" in workbook.sheetnames:
            _apply_row_status_styles(workbook["Transferencias"])
            _apply_currency_format(
                workbook["Transferencias"], ("Neto", "Transferencia", "Devengado", "IBC")
            )
        if "Seguridad_Social" in workbook.sheetnames:
            _apply_row_status_styles(workbook["Seguridad_Social"])
            _apply_currency_format(workbook["Seguridad_Social"], ("Devengado", "IBC"))

    output.seek(0)
    return output.getvalue()
