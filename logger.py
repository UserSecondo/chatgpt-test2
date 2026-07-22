"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : logger.py
Versión  : 5.1

Configuración centralizada del logger del proyecto.

Responsabilidades
-----------------
- Crear el logger principal.
- Registrar mensajes en consola y archivo.
- Mantener una única configuración de logging.

No realiza:
- Reglas de negocio.
- Lectura de archivos.
- Escritura de diccionarios.
===============================================================================
"""

from __future__ import annotations

import logging
from pathlib import Path

from metadata import DIRECTORIES, FILE_EXTENSIONS, MESSAGES, OUTPUT_FILENAMES

_LOGGER_NAME = "BDD_GEO_DICTIONARY3"


def get_logger() -> logging.Logger:
    """
    Obtiene el logger principal del proyecto.

    Returns
    -------
    logging.Logger
        Logger configurado para consola y archivo.
    """

    logger = logging.getLogger(_LOGGER_NAME)

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    logger.propagate = False

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    log_dir = Path(DIRECTORIES["logs"])
    log_dir.mkdir(parents=True, exist_ok=True)

    log_path = log_dir / OUTPUT_FILENAMES["log"]
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.info(MESSAGES["project_start"])

    return logger
