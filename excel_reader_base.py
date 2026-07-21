"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : excel_reader_base.py
Versión  : 5.1

Construye el modelo ProjectMetadata a partir de los archivos Excel.

Responsabilidades
-----------------
- Leer las diferentes fuentes de información.
- Construir el modelo de dominio.
- Mantener índices internos para búsquedas rápidas.

No realiza:
- Resolución de conflictos entre fuentes.
- Aplicación de reglas de negocio.
- Escritura del diccionario de datos.
===============================================================================
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import Config

from metadata import (
    COLUMN_MAPS,
    INPUT_FILENAMES,
    REQUIRED_COLUMNS,
    SHEETS,
)

from models import (
    FieldInfo,
    ProjectMetadata,
    SchemaInfo,
    TableInfo,
)


class ExcelReaderBase:
    """
    Construye el modelo ProjectMetadata leyendo todas
    las fuentes Excel del proyecto.
    """

    ###########################################################################
    # Constructor
    ###########################################################################

    def __init__(self, config: Config):

        self.config = config

        self.project = ProjectMetadata()

        #
        # Índices internos
        #

        self._schemas: Dict[str, SchemaInfo] = {}

        self._tables: Dict[Tuple[str, str], TableInfo] = {}

        self._fields: Dict[
            Tuple[str, str, str],
            FieldInfo,
        ] = {}

    ###########################################################################
    # API pública
    ###########################################################################

    def read(self) -> ProjectMetadata:
        """
        Construye el modelo completo del proyecto.
        """

        self._load_oracle()

        self._load_inventory()

        self._load_catalog()

        self._load_esri()

        self._load_mgn()

        self._update_summary()

        self._validate_project()

        return self.project

    ###########################################################################
    # Apertura de archivos
    ###########################################################################

    def _open_workbook(
        self,
        filename: str,
    ) -> Workbook:
        """
        Abre un archivo Excel en modo lectura.
        """

        path = self.config.input_file(filename)

        if not path.exists():

            raise FileNotFoundError(

                f"No existe el archivo:\n{path}"

            )

        return load_workbook(

            filename=path,

            data_only=True,

            read_only=True,

        )

    ###########################################################################
    # Utilidades de lectura
    ###########################################################################

    @staticmethod
    def _worksheet_headers(
        sheet: Worksheet,
    ) -> list[str]:
        """
        Obtiene los encabezados de una hoja.
        """

        headers = []

        for cell in sheet[1]:

            value = cell.value

            if value is None:

                headers.append("")

            else:

                headers.append(

                    str(value).strip()

                )

        return headers

    ###########################################################################

    @staticmethod
    def _normalize_text(
        value: Any,
    ) -> str:
        """
        Normaliza texto libre.

        No modifica mayúsculas/minúsculas.
        """

        if value is None:

            return ""

        return str(value).strip()

    ###########################################################################

    @staticmethod
    def _normalize_identifier(
        value: Any,
    ) -> str:
        """
        Normaliza identificadores.

        Se utiliza para:

        - esquema
        - tabla
        - campo
        """

        if value is None:

            return ""

        return str(value).strip().upper()

    ###########################################################################

    def _iter_rows(
        self,
        sheet: Worksheet,
    ) -> Iterable[Dict[str, str]]:
        """
        Convierte una hoja Excel en un iterador
        de diccionarios.
        """

        headers = self._worksheet_headers(sheet)

        for row in sheet.iter_rows(

            min_row=2,

            values_only=True,

        ):

            yield {

                headers[i]:

                self._normalize_text(value)

                for i, value in enumerate(row)

            }

    ###########################################################################
    # Validación
    ###########################################################################

    @staticmethod
    def _validate_columns(
        headers: list[str],
        required: tuple,
    ) -> None:
        """
        Verifica que existan todas las columnas
        obligatorias.
        """

        headers_upper = {

            h.upper()

            for h in headers

        }

        missing = [

            col

            for col in required

            if col.upper() not in headers_upper

        ]

        if missing:

            raise ValueError(

                "Columnas obligatorias no encontradas: "

                + ", ".join(missing)

            )

    ###########################################################################
    # Conversión al modelo lógico
    ###########################################################################

    @staticmethod
    def _logical_row(
        row: Dict[str, str],
        mapping: Dict[str, str],
    ) -> Dict[str, str]:
        """
        Convierte una fila del Excel al modelo lógico
        definido en metadata.py.
        """

        logical = {}

        for excel_name, logical_name in mapping.items():

            logical[logical_name] = row.get(
                excel_name,
                "",
            )

        return logical

    ###########################################################################
    # Lectura genérica de hojas
    ###########################################################################

    def _read_sheet(
        self,
        workbook: Workbook,
        sheet_name: str,
        mapping_name: str,
    ) -> Iterable[Dict[str, str]]:
        """
        Lee una hoja Excel y devuelve las filas
        convertidas al modelo lógico.
        """

        try:

            sheet = workbook[sheet_name]

        except KeyError as ex:

            raise ValueError(
                f"No existe la hoja '{sheet_name}'."
            ) from ex

        headers = self._worksheet_headers(sheet)

        self._validate_columns(
            headers,
            REQUIRED_COLUMNS[mapping_name],
        )

        mapping = COLUMN_MAPS[mapping_name]

        for row in self._iter_rows(sheet):

            logical = self._logical_row(
                row,
                mapping,
            )

            #
            # Normalización de identificadores
            #

            if "schema" in logical:
                logical["schema"] = self._normalize_identifier(
                    logical["schema"]
                )

            if "table" in logical:
                logical["table"] = self._normalize_identifier(
                    logical["table"]
                )

            if "field" in logical:
                logical["field"] = self._normalize_identifier(
                    logical["field"]
                )

            yield logical

    ###########################################################################
    # Índices internos
    ###########################################################################

    def _get_schema(
        self,
        schema_name: str,
    ) -> SchemaInfo:
        """
        Obtiene un esquema existente o lo crea.
        """

        schema = self._schemas.get(schema_name)

        if schema is None:

            schema = SchemaInfo(
                schema_name=schema_name
            )

            self.project.schemas.append(schema)

            self._schemas[schema_name] = schema

        return schema

    ###########################################################################

    def _get_table(
        self,
        schema_name: str,
        table_name: str,
    ) -> TableInfo:
        """
        Obtiene una tabla existente o la crea.
        """

        key = (
            schema_name,
            table_name,
        )

        table = self._tables.get(key)

        if table is None:

            schema = self._get_schema(schema_name)

            table = TableInfo(
                schema_name=schema_name,
                table_name=table_name,
            )

            schema.tables.append(table)

            self._tables[key] = table

        return table

    ###########################################################################

    def _get_field(
        self,
        schema_name: str,
        table_name: str,
        field_name: str,
    ) -> FieldInfo:
        """
        Obtiene un campo existente o lo crea.
        """

        key = (
            schema_name,
            table_name,
            field_name,
        )

        field = self._fields.get(key)

        if field is None:

            table = self._get_table(
                schema_name,
                table_name,
            )

            field = FieldInfo(
                field_name=field_name
            )

            table.fields.append(field)

            self._fields[key] = field

        return field

    ###########################################################################
    # Estadísticas
    ###########################################################################

    def _update_summary(self) -> None:
        """
        Actualiza el resumen de la ejecución.
        """

        summary = self.project.summary

        summary.processed_schemas = len(
            self.project.schemas
        )

        summary.processed_tables = sum(
            len(schema.tables)
            for schema in self.project.schemas
        )

        summary.processed_fields = sum(
            len(table.fields)
            for schema in self.project.schemas
            for table in schema.tables
        )
