"""Trazas de depuracion, activables con VALIDATION_DEBUG=0/1."""
from __future__ import annotations

import os

DEBUG = os.environ.get("VALIDATION_DEBUG", "1") == "1"


def debug(mensaje: str) -> None:
    if DEBUG:
        print(f"[DEBUG][facturacion] {mensaje}")
