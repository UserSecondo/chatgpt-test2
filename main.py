"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : main.py
Autor    : Nelson David Martínez
Versión  : 4.0

Descripción
-----------
Punto de entrada del proyecto.

Responsabilidades
-----------------
1. Inicializar la aplicación.
2. Cargar la configuración.
3. Leer todas las fuentes de información.
4. Resolver los metadatos aplicando las reglas de negocio.
5. Generar los diccionarios de datos.
6. Registrar el resultado de la ejecución.

IMPORTANTE
-----------
Este módulo NO contiene:

- Reglas de negocio.
- Lectura de archivos Excel.
- Escritura de archivos Excel.
- Transformaciones de datos.

Únicamente coordina la ejecución del proyecto.
===============================================================================
"""

from __future__ import annotations

import sys
from datetime import datetime

from config import Config
from excel_reader import ExcelReader
from metadata_resolver import MetadataResolver
from excel_writer import ExcelWriter
from logger import get_logger


# =============================================================================
# Función principal
# =============================================================================

def main() -> int:
    """
    Punto de entrada de la aplicación.

    Returns
    -------
    int
        0 : Ejecución exitosa.
        1 : Error durante la ejecución.
    """

    logger = get_logger()

    start_time = datetime.now()

    logger.info("=" * 80)
    logger.info("BDD_GEO_DICTIONARY3")
    logger.info("Generador de Diccionarios de Datos")
    logger.info("=" * 80)

    try:

        # ---------------------------------------------------------------------
        # Configuración
        # ---------------------------------------------------------------------

        logger.info("Cargando configuración...")

        config = Config()

        logger.info("Configuración cargada correctamente.")

        # ---------------------------------------------------------------------
        # Lectura de metadatos
        # ---------------------------------------------------------------------

        logger.info("Leyendo fuentes de información...")

        reader = ExcelReader(config)

        project_metadata = reader.read()

        logger.info("Lectura finalizada correctamente.")

        # ---------------------------------------------------------------------
        # Aplicación de reglas de negocio
        # ---------------------------------------------------------------------

        logger.info("Resolviendo metadatos...")

        resolver = MetadataResolver()

        resolved_metadata = resolver.resolve(project_metadata)

        logger.info("Metadatos resueltos correctamente.")

        # ---------------------------------------------------------------------
        # Generación de diccionarios
        # ---------------------------------------------------------------------

        logger.info("Generando diccionarios de datos...")

        writer = ExcelWriter(config)

        writer.generate(resolved_metadata)

        logger.info("Diccionarios generados correctamente.")

        # ---------------------------------------------------------------------
        # Resumen
        # ---------------------------------------------------------------------

        elapsed = datetime.now() - start_time

        logger.info("-" * 80)
        logger.info("Proceso finalizado correctamente.")
        logger.info(f"Tiempo total de ejecución: {elapsed}")
        logger.info("-" * 80)

        return 0

    except Exception:

        logger.exception("Se produjo un error durante la ejecución.")

        return 1


# =============================================================================
# Inicio del programa
# =============================================================================

if __name__ == "__main__":

    sys.exit(main())