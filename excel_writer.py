"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : excel_writer.py
Versión  : 7.0

Genera el diccionario de datos institucional a partir del modelo resuelto.

Formato de salida
------------------
- Un archivo Excel POR ESQUEMA documentado.
- Dentro de cada archivo, una hoja POR TABLA.
- Cada hoja respeta la plantilla institucional
  (Plantilla_Diccionario_Datos_DBGEODIG_AJUSTADO.xlsx):

    Esquema:
    Tabla:
    Descripción de la tabla:
    Responsable:
    Fecha de documentación:

    N° | Campo | Tipo de dato (Oracle) | Longitud/Precisión (Oracle)
       | Tipo de dato (ESRI) | Longitud/Precisión (ESRI)
       | PK | FK | Nulo | Descripción

Notas
-----
- PK y FK se dejan en blanco: ninguna fuente actual provee esa
  información (ver conversación con el usuario).
- No realiza lectura de archivos de entrada, ni reglas de negocio,
  ni resolución de conflictos entre fuentes.
===============================================================================
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import Config
from models import FieldInfo, ProjectMetadata, SchemaInfo, TableInfo

_TEMPLATE_SHEET_NAME = "Diccionario_Datos"

_FIRST_DATA_ROW = 8

_INVALID_SHEET_CHARS = r"[\[\]:\*\?/\\]"

_INVALID_FILENAME_CHARS = r'[\\/:*?"<>|]'


class ExcelWriter:
    """
    Genera el diccionario de datos institucional: un archivo por
    esquema, con una hoja por tabla, siguiendo la plantilla real.
    """

    def __init__(self, config: Config):
        self.config = config

    def generate(self, project: ProjectMetadata) -> list[Path]:
        """
        Genera un archivo Excel por cada esquema documentado.

        Returns
        -------
        list[Path]
            Rutas de los archivos generados.
        """

        template_path = self.config.template_file

        generated_paths: list[Path] = []

        for schema in project.schemas:

            if not schema.tables:
                continue

            output_path = self._generate_schema_workbook(
                schema,
                template_path,
            )

            generated_paths.append(output_path)

            project.summary.generated_files += 1

        project.summary.end_time = datetime.now()

        if project.summary.start_time and project.summary.end_time:
            project.summary.execution_time = (
                project.summary.end_time - project.summary.start_time
            )

        return generated_paths

    ###########################################################################
    # Generación por esquema
    ###########################################################################

    def _generate_schema_workbook(
        self,
        schema: SchemaInfo,
        template_path: Path,
    ) -> Path:
        """
        Genera el archivo Excel de un esquema, con una hoja por tabla.
        """

        workbook = load_workbook(template_path)

        master_sheet = workbook[_TEMPLATE_SHEET_NAME]

        template_last_row = master_sheet.max_row

        used_sheet_names: set[str] = set()

        for table in schema.tables:

            sheet_name = self._unique_sheet_name(
                table.table_name,
                used_sheet_names,
            )

            sheet = workbook.copy_worksheet(master_sheet)
            sheet.title = sheet_name

            self._write_table_sheet(sheet, schema, table, template_last_row)

        # La hoja maestra y las notas de uso son insumos internos
        # de la plantilla; no forman parte del diccionario final.

        del workbook[_TEMPLATE_SHEET_NAME]

        if "Hoja2" in workbook.sheetnames:
            del workbook["Hoja2"]

        output_path = self.config.output_file(
            self._schema_filename(schema.schema_name)
        )

        workbook.save(output_path)
        workbook.close()

        return output_path

    ###########################################################################
    # Generación por tabla
    ###########################################################################

    def _write_table_sheet(
        self,
        sheet: Worksheet,
        schema: SchemaInfo,
        table: TableInfo,
        template_last_row: int,
    ) -> None:
        """
        Llena una hoja (copia de la plantilla) con la
        información de una tabla.

        El número de filas de campos queda acorde a la cantidad
        real de campos de la tabla: se eliminan las filas de
        ejemplo sobrantes cuando hay menos campos que en la
        plantilla, y se agregan filas adicionales cuando hay más.
        """

        sheet["B2"] = schema.schema_name
        sheet["B3"] = table.table_name
        sheet["B4"] = table.description
        sheet["B5"] = schema.responsible
        sheet["B6"] = datetime.now().strftime("%Y-%m-%d")

        self._clear_placeholder_rows(sheet, template_last_row)

        row_index = _FIRST_DATA_ROW

        for position, field in enumerate(table.fields, start=1):
            self._write_field_row(sheet, row_index, position, field)
            row_index += 1

        last_written_row = row_index - 1

        if last_written_row < template_last_row:

            # Menos campos que filas de ejemplo en la plantilla:
            # se eliminan las filas sobrantes.

            sheet.delete_rows(
                last_written_row + 1,
                template_last_row - last_written_row,
            )

    @staticmethod
    def _write_field_row(
        sheet: Worksheet,
        row_index: int,
        position: int,
        field: FieldInfo,
    ) -> None:
        """
        Escribe una fila de campo siguiendo el orden de columnas
        de la plantilla institucional.
        """

        values = [
            position,
            field.field_name,
            field.oracle_type,
            field.oracle_length,
            field.esri_type,
            field.esri_length,
            "",  # PK: sin fuente disponible actualmente.
            "",  # FK: sin fuente disponible actualmente.
            field.nullable,
            field.description,
        ]

        for col_idx, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=col_idx, value=value)

    @staticmethod
    def _clear_placeholder_rows(
        sheet: Worksheet,
        template_last_row: int,
    ) -> None:
        """
        La plantilla trae filas de ejemplo numeradas bajo el
        encabezado (hasta la fila `template_last_row`). Se
        limpian antes de escribir los datos reales para no
        dejar residuos de la plantilla.
        """

        for row in range(_FIRST_DATA_ROW, template_last_row + 1):
            for col in range(1, 11):
                sheet.cell(row=row, column=col, value=None)

    ###########################################################################
    # Nombres de hoja y de archivo
    ###########################################################################

    @staticmethod
    def _unique_sheet_name(
        table_name: str,
        used_names: set[str],
    ) -> str:
        """
        Genera un nombre de hoja válido para Excel (máximo 31
        caracteres, sin caracteres inválidos) y único dentro
        del libro.
        """

        base_name = re.sub(_INVALID_SHEET_CHARS, "_", table_name)[:31]

        candidate = base_name
        suffix = 1

        while candidate in used_names:

            suffix_text = f"_{suffix}"
            candidate = base_name[: 31 - len(suffix_text)] + suffix_text
            suffix += 1

        used_names.add(candidate)

        return candidate

    @staticmethod
    def _schema_filename(schema_name: str) -> str:
        """
        Genera el nombre de archivo del diccionario de un esquema.
        """

        safe_name = re.sub(_INVALID_FILENAME_CHARS, "_", schema_name)

        return f"Diccionario_Datos_{safe_name}.xlsx"
