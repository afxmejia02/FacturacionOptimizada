"""Trazas de depuracion, activables con VALIDATION_DEBUG=0/1."""
from __future__ import annotations

import os


def log(mensaje: str) -> None:
    if os.environ.get("VALIDATION_DEBUG", "1") != "0":
        print(f"[DEBUG][nomina] {mensaje}")
