"""Presentation helpers: HTML rendering and value formatting for the web UI.

These functions are pure (no Streamlit dependency) so they can be unit-tested
and reused by the entry point in ``app.py``.
"""
from __future__ import annotations

import pandas as pd


def build_colored_table(df_display: pd.DataFrame) -> str:
    """Render a DataFrame as an HTML table coloured by the ``Estado`` column.

    Rows whose ``Estado`` is ``OK`` are shown in green; anything else in red.
    """
    headers = list(df_display.columns)
    parts = [
        '<table style="width:100%; border-collapse:collapse; font-family:Arial,sans-serif;">',
        '<thead><tr>',
    ]
    for header in headers:
        parts.append(
            '<th style="border:1px solid #ddd; padding:8px; text-align:left; background:#f4f4f4;">'
            f"{header}</th>"
        )
    parts.append('</tr></thead>')
    parts.append('<tbody>')

    for _, row in df_display.iterrows():
        estado = str(row.get("Estado", "")).strip().lower()
        if estado == "ok":
            row_style = "background-color:#d4edda; color:#155724;"
        else:
            row_style = "background-color:#f8d7da; color:#721c24;"
        parts.append(f'<tr style="{row_style}">')
        for header in headers:
            value = row[header] if pd.notna(row[header]) else ""
            parts.append(
                '<td style="border:1px solid #ddd; padding:8px; vertical-align:top;">'
                f"{value}</td>"
            )
        parts.append('</tr>')

    parts.append('</tbody></table>')
    return "".join(parts)


def _cell_text(value) -> str:
    """Render a single comparison value, blanking out NaN/None."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def build_mano_obra_table(
    df_display: pd.DataFrame,
    source_labels: tuple[str, str] = ("Informe", "ODS"),
    estado_col: str = "Estado revisión",
) -> str:
    """Render the mano-de-obra comparison, colouring only the inconsistent cells.

    Comparison columns hold lists: a single element means both sources agree
    (shown plain); two elements means they differ (shown as ``Informe / ODS`` on
    a red cell). The rest of the row keeps a neutral background, so only the
    column with the inconsistency is highlighted.
    """
    headers = list(df_display.columns)
    base_td = "border:1px solid #ddd; padding:8px; vertical-align:top;"
    diff_td = base_td + " background-color:#f8d7da; color:#721c24; font-weight:bold;"
    ok_td = base_td + " background-color:#d4edda; color:#155724;"

    parts = [
        '<table style="width:100%; border-collapse:collapse; font-family:Arial,sans-serif;">',
        '<thead><tr>',
    ]
    for header in headers:
        parts.append(
            '<th style="border:1px solid #ddd; padding:8px; text-align:left; background:#f4f4f4;">'
            f"{header}</th>"
        )
    parts.append('</tr></thead><tbody>')

    for _, row in df_display.iterrows():
        parts.append('<tr>')
        for header in headers:
            value = row[header]
            if isinstance(value, (list, tuple)):
                if len(value) <= 1:
                    text = _cell_text(value[0]) if value else ""
                    style = base_td
                else:
                    text = (
                        f"<b>{source_labels[0]}:</b> {_cell_text(value[0])}<br>"
                        f"<b>{source_labels[1]}:</b> {_cell_text(value[1])}"
                    )
                    style = diff_td
            elif header == estado_col:
                text = _cell_text(value)
                style = ok_td if text.strip().lower() == "ok" else diff_td
            else:
                text = _cell_text(value)
                style = base_td
            parts.append(f'<td style="{style}">{text}</td>')
        parts.append('</tr>')

    parts.append('</tbody></table>')
    return "".join(parts)


def format_count(value):
    """Normalise a count to int when it has no fractional part, else float."""
    if value is None:
        return 0
    try:
        import numpy as _np
    except Exception:
        _np = None

    if _np is not None and isinstance(value, _np.generic):
        value = value.item()
    try:
        numeric = round(float(value), 6)  # absorbe el ruido de coma flotante (483.46000000000004 -> 483.46)
        if numeric.is_integer():
            return int(numeric)
        return numeric
    except Exception:
        return value


def normalize_list_like(value) -> str:
    """Turn a scalar or list-like of numbers into a space-separated string."""
    if value is None:
        return ""
    try:
        import numpy as _np
    except Exception:
        _np = None

    if isinstance(value, (list, tuple, set)):
        parts = []
        for item in value:
            try:
                if _np is not None and isinstance(item, _np.generic):
                    item = int(item)
                parts.append(str(int(item)))
            except Exception:
                parts.append(str(item))
        return " ".join(parts)

    try:
        if _np is not None and isinstance(value, _np.generic):
            return str(int(value))
        if isinstance(value, (int, float)):
            return str(int(value))
    except Exception:
        pass

    return str(value)


def format_dataframe(df_in: pd.DataFrame, formatter) -> pd.DataFrame:
    """Apply ``formatter`` to the list-like reconciliation columns for display."""
    if df_in is None or df_in.empty:
        return df_in
    df_tmp = df_in.copy()
    list_cols = [
        col
        for col in ["Neto_desprendibles", "Valores_transferencia", "Devengado", "IBC"]
        if col in df_tmp.columns
    ]

    for column in list_cols:
        def _fmt_cell(cell):
            normalized = normalize_list_like(cell)
            try:
                return formatter(normalized)
            except Exception:
                return normalized

        df_tmp[column] = df_tmp[column].apply(_fmt_cell)
    return df_tmp
