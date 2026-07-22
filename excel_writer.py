"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : excel_writer.py
Versión  : 5.1

Genera el diccionario de datos en formato Excel a partir del modelo resuelto.

Responsabilidades
-----------------
- Crear el libro de salida.
- Volcar el metadata resuelto a hojas Excel.
- Guardar el archivo final del diccionario.

No realiza:
- Lectura de archivos Excel de entrada.
- Reglas de negocio.
- Resolución de conflictos.
===============================================================================
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from config import Config
from metadata import OUTPUT_FILENAMES
from models import FieldInfo, ProjectMetadata, SchemaInfo, TableInfo


class ExcelWriter:
    """
    Genera el archivo Excel final del diccionario de datos.
    """

    def __init__(self, config: Config):
        self.config = config

    def generate(self, project: ProjectMetadata) -> Path:
        """
        Genera el archivo Excel del diccionario de datos.

        Parameters
        ----------
        project : ProjectMetadata
            Modelo resuelto del proyecto.

        Returns
        -------
        Path
            Ruta del archivo generado.
        """

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Diccionario"

        headers = [
            "Esquema",
            "Tabla",
            "Campo",
            "Tipo",
            "Longitud",
            "Nullable",
            "Dominio",
            "Alias",
            "Descripción",
            "Observaciones",
            "Responsable esquema",
            "Descripción tabla",
            "Fuente descripción campo",
        ]

        self._write_headers(sheet, headers)

        row_index = 2
        for schema in project.schemas:
            row_index = self._write_schema(sheet, schema, row_index)

        output_path = self.config.output_file(OUTPUT_FILENAMES["dictionary"])
        workbook.save(output_path)
        workbook.close()

        project.summary.generated_files += 1
        project.summary.end_time = datetime.now()
        if project.summary.start_time and project.summary.end_time:
            project.summary.execution_time = (
                project.summary.end_time - project.summary.start_time
            )

        return output_path

    def _write_headers(self, sheet, headers: list[str]) -> None:
        """Escribe los encabezados de la hoja principal."""
        for col_idx, header in enumerate(headers, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def _write_schema(self, sheet, schema: SchemaInfo, start_row: int) -> int:
        """Escribe un esquema y sus tablas/campos."""
        current_row = start_row
        for table in schema.tables:
            current_row = self._write_table(sheet, schema, table, current_row)
        return current_row

    def _write_table(
        self,
        sheet,
        schema: SchemaInfo,
        table: TableInfo,
        start_row: int,
    ) -> int:
        """Escribe una tabla y sus campos."""
        current_row = start_row

        if not table.fields:
            sheet.cell(row=current_row, column=1, value=schema.schema_name)
            sheet.cell(row=current_row, column=2, value=table.table_name)
            sheet.cell(row=current_row, column=12, value=table.description)
            current_row += 1
            return current_row

        for field in table.fields:
            self._write_field(sheet, schema, table, field, current_row)
            current_row += 1

        return current_row

    def _write_field(
        self,
        sheet,
        schema: SchemaInfo,
        table: TableInfo,
        field: FieldInfo,
        row: int,
    ) -> None:
        """Escribe una fila de campo."""
        values = [
            schema.schema_name,
            table.table_name,
            field.field_name,
            field.data_type,
            field.length,
            field.nullable,
            field.domain,
            field.alias,
            field.description,
            field.observations,
            schema.responsible,
            table.description,
            field.description_source,
        ]

        for col_idx, value in enumerate(values, start=1):
            sheet.cell(row=row, column=col_idx, value=value)
