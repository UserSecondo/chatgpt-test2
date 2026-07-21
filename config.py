"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : config.py
Autor    : Nelson David Martínez
Versión  : 5.0

Descripción
-----------
Configuración general del proyecto.

Responsabilidades
-----------------
- Definir la estructura de carpetas.
- Centralizar las rutas del proyecto.
- Validar la existencia de los archivos de entrada.
- Crear automáticamente la estructura de directorios.

IMPORTANTE
-----------
Este módulo NO contiene:

- Reglas de negocio.
- Lectura de archivos Excel.
- Escritura de archivos.
- Transformación de datos.
===============================================================================
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterator


class Config:
    """
    Configuración general del proyecto.
    """

    # =========================================================================
    # Constructor
    # =========================================================================

    def __init__(self) -> None:

        # ---------------------------------------------------------------------
        # Directorio raíz
        # ---------------------------------------------------------------------

        self.BASE_DIR: Path = Path(__file__).resolve().parent

        # ---------------------------------------------------------------------
        # Directorios del proyecto
        # ---------------------------------------------------------------------

        self.INPUT_DIR: Path = self.BASE_DIR / "input"

        self.OUTPUT_DIR: Path = self.BASE_DIR / "output"

        self.LOG_DIR: Path = self.BASE_DIR / "logs"

        self.DOCS_DIR: Path = self.BASE_DIR / "docs"

        self.TEST_DIR: Path = self.BASE_DIR / "tests"

        # ---------------------------------------------------------------------
        # Archivos de entrada
        # ---------------------------------------------------------------------

        self.INPUT_FILES: Dict[str, Path] = {

            "oracle":
                self.INPUT_DIR / "ESQUEMAS_ORACLE_DBGEODIG.xlsx",

            "inventory":
                self.INPUT_DIR / "inventario_oracle_DBGEODIG_20260708.xlsx",

            "catalog":
                self.INPUT_DIR / "13052026_CatalogoBDxEsquemas.xlsx",

            "esri":
                self.INPUT_DIR / "Inventario_Vectorial_DBGEODIG_ESRI_20260707.xlsx",

            "mgn":
                self.INPUT_DIR / "Diccionario_Datos_MGN_2005.xlsx",

            "template":
                self.INPUT_DIR / "Plantilla_Diccionario_Datos_DBGEODIG_AJUSTADO.xlsx"

        }

        # ---------------------------------------------------------------------
        # Inicialización
        # ---------------------------------------------------------------------

        self.create_directories()

        self.validate()

    # =========================================================================
    # Directorios
    # =========================================================================

    def create_directories(self) -> None:
        """
        Crea automáticamente la estructura de carpetas del proyecto.
        """

        for directory in (

            self.INPUT_DIR,
            self.OUTPUT_DIR,
            self.LOG_DIR,
            self.DOCS_DIR,
            self.TEST_DIR,

        ):

            directory.mkdir(parents=True, exist_ok=True)

    # =========================================================================
    # Validación
    # =========================================================================

    def validate(self) -> None:
        """
        Verifica la existencia de todos los archivos obligatorios.
        """

        missing = [

            path.name

            for path in self.INPUT_FILES.values()

            if not path.exists()

        ]

        if missing:

            message = (
                "\n"
                "No se encontraron los siguientes archivos de entrada:\n\n"
                + "\n".join(f" • {file}" for file in missing)
            )

            raise FileNotFoundError(message)

    # =========================================================================
    # Métodos públicos
    # =========================================================================

    def input_file(self, key: str) -> Path:
        """
        Retorna la ruta de un archivo de entrada.
        """

        return self.INPUT_FILES[key]

    # -------------------------------------------------------------------------

    def output_file(self, filename: str) -> Path:
        """
        Construye la ruta de un archivo de salida.
        """

        return self.OUTPUT_DIR / filename

    # -------------------------------------------------------------------------

    def log_file(self, filename: str) -> Path:
        """
        Construye la ruta de un archivo de log.
        """

        return self.LOG_DIR / filename

    # -------------------------------------------------------------------------

    @property
    def template_file(self) -> Path:
        """
        Retorna la plantilla oficial del diccionario de datos.
        """

        return self.INPUT_FILES["template"]

    # -------------------------------------------------------------------------

    def exists(self, key: str) -> bool:
        """
        Verifica si existe un archivo de entrada.
        """

        return self.INPUT_FILES[key].exists()

    # -------------------------------------------------------------------------

    def all_input_files(self) -> Iterator[tuple[str, Path]]:
        """
        Permite recorrer todos los archivos de entrada.

        Ejemplo
        -------
        for name, path in config.all_input_files():
            ...
        """

        return iter(self.INPUT_FILES.items())

    # =========================================================================

    def __str__(self) -> str:

        lines = [

            "",
            "=" * 80,
            "CONFIGURACIÓN DEL PROYECTO",
            "=" * 80,
            f"Base       : {self.BASE_DIR}",
            f"Input      : {self.INPUT_DIR}",
            f"Output     : {self.OUTPUT_DIR}",
            f"Logs       : {self.LOG_DIR}",
            f"Docs       : {self.DOCS_DIR}",
            f"Tests      : {self.TEST_DIR}",
            "=" * 80,

        ]

        return "\n".join(lines)