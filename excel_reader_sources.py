"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : excel_reader_sources.py
Versión  : 5.1

Lectura de las fuentes ESRI, MGN y validación del modelo.

Responsabilidades
-----------------
- Enriquecer campos con metadata ESRI.
- Enriquecer campos con metadata MGN.
- Validar la consistencia del modelo cargado.

No realiza:
- Resolución de conflictos entre fuentes.
- Aplicación de reglas de negocio.
- Escritura del diccionario de datos.
===============================================================================
"""

from __future__ import annotations

from openpyxl.workbook.workbook import Workbook

from metadata import SHEETS


class ExcelReaderSources:
    """
    Implementa la lectura de ESRI, MGN y validaciones.
    """

    ###########################################################################
    # ESRI
    ###########################################################################

    def _load_esri(self) -> None:
        """
        Procesa el archivo de metadatos ESRI.
        """

        workbook = self._open_workbook("esri")

        try:
            self._read_esri_fields(workbook)
        finally:
            workbook.close()

    def _read_esri_fields(self, workbook: Workbook) -> None:
        """
        Enriquece los campos utilizando la metadata ESRI.

        Esta fuente NO crea nuevos objetos del modelo.
        Únicamente complementa la información de los
        campos ya existentes.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["esri"],
            "esri",
        ):
            schema_name = row.get("schema", "")
            table_name = row.get("table", "")
            field_name = row.get("field", "")

            if not schema_name or not table_name or not field_name:
                continue

            if not self._is_documented(schema_name):
                continue

            field = self._get_field(schema_name, table_name, field_name)

            esri_type = row.get("type", "")
            if esri_type:
                field.esri_type = esri_type

            esri_length = row.get("length", "")
            if esri_length:
                field.esri_length = str(esri_length)

            self._assign_by_priority(
                field,
                "data_type",
                "data_type_source",
                "field_description",
                esri_type,
                "esri",
            )

            self._set_if_value(
                field,
                "alias",
                row.get("alias", ""),
            )

            self._set_if_value(
                field,
                "domain",
                row.get("domain", ""),
            )

            self._assign_by_priority(
                field,
                "description",
                "description_source",
                "field_description",
                row.get("description", ""),
                "esri",
            )

            self._assign_if_empty(
                field,
                "observations",
                row.get("observations", ""),
            )

    ###########################################################################
    # MGN
    ###########################################################################

    def _load_mgn(self) -> None:
        """
        Procesa el archivo de metadatos MGN.
        """

        workbook = self._open_workbook("mgn")

        try:
            self._read_mgn(workbook)
        finally:
            workbook.close()

    def _read_mgn(self, workbook: Workbook) -> None:
        """
        Enriquece los campos utilizando la metadata MGN.

        A diferencia de Oracle o ESRI, el diccionario MGN documenta
        campos estándar únicamente por NOMBRE (por ejemplo DPTO_CCDGO,
        MPIO_CCDGO), sin indicar a qué esquema o tabla pertenecen.
        Además cada hoja contiene varios bloques de datos, cada uno
        con su propio encabezado 'CAMPOS'/'DESCRIPCIÓN' repetido.

        Por eso la enriquece de forma global: cualquier campo del
        proyecto cuyo nombre coincida recibe la descripción MGN
        correspondiente, sin importar en qué esquema o tabla esté.
        """

        mgn_descriptions: dict[str, str] = {}

        for sheet_name in SHEETS["mgn"]:

            sheet = workbook[sheet_name]

            for field_name, description in self._iter_mgn_blocks(sheet):

                if not field_name or not description:
                    continue

                mgn_descriptions.setdefault(field_name, description)

        if not mgn_descriptions:
            return

        for schema in self.project.schemas:
            for table in schema.tables:
                for field in table.fields:

                    description = mgn_descriptions.get(field.field_name)

                    if description is None:
                        continue

                    self._assign_by_priority(
                        field,
                        "description",
                        "description_source",
                        "field_description",
                        description,
                        "mgn",
                    )

    @staticmethod
    def _iter_mgn_blocks(sheet):
        """
        Recorre una hoja MGN detectando cada bloque de datos.

        Cada hoja documenta varias feature classes seguidas, y cada
        una repite su propio encabezado 'CAMPOS' / 'DESCRIPCIÓN' en
        una posición de columna que puede variar de una hoja a otra.
        Esta función localiza cada encabezado y produce las filas de
        datos que le siguen, ignorando los encabezados repetidos.
        """

        field_col = None
        description_col = None

        for row in sheet.iter_rows(values_only=True):

            header_positions = {
                str(value).strip().upper(): idx
                for idx, value in enumerate(row)
                if value is not None
            }

            if "CAMPOS" in header_positions and "DESCRIPCIÓN" in header_positions:
                field_col = header_positions["CAMPOS"]
                description_col = header_positions["DESCRIPCIÓN"]
                continue

            if field_col is None or description_col is None:
                continue

            if field_col >= len(row) or description_col >= len(row):
                continue

            field_name = row[field_col]

            if field_name is None:
                continue

            description = row[description_col]

            yield (
                str(field_name).strip().upper(),
                str(description).strip() if description is not None else "",
            )

    ###########################################################################
    # Validación del modelo
    ###########################################################################

    def _validate_project(self) -> None:
        """
        Ejecuta todas las validaciones del modelo cargado.
        """

        self._validate_schemas()
        self._validate_tables()
        self._validate_fields()

    def _validate_schemas(self) -> None:
        """
        Valida la información de los esquemas.
        """

        for schema in self.project.schemas:
            if not schema.tables:
                self.project.warnings.append(
                    f"El esquema '{schema.schema_name}' no contiene tablas."
                )

    def _validate_tables(self) -> None:
        """
        Valida la información de las tablas.
        """

        for schema in self.project.schemas:
            for table in schema.tables:
                if not table.fields:
                    self.project.warnings.append(
                        f"La tabla '{table.table_name}' no contiene campos."
                    )

    def _validate_fields(self) -> None:
        """
        Valida los campos del modelo.
        """

        for schema in self.project.schemas:
            for table in schema.tables:
                for field in table.fields:
                    if not field.data_type:
                        self.project.warnings.append(
                            f"{schema.schema_name}.{table.table_name}.{field.field_name} no tiene tipo de dato."
                        )

                    if not field.description:
                        self.project.warnings.append(
                            f"{schema.schema_name}.{table.table_name}.{field.field_name} no tiene descripción."
                        )
