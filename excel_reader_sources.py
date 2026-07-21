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

from metadata import INPUT_FILENAMES, SHEETS


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

        workbook = self._open_workbook(
            INPUT_FILENAMES["esri"]
        )

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

            field = self._get_field(schema_name, table_name, field_name)

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

            self._set_if_value(
                field,
                "description",
                row.get("description", ""),
                "description_source",
                "esri",
            )

            self._set_if_value(
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

        workbook = self._open_workbook(
            INPUT_FILENAMES["mgn"]
        )

        try:
            self._read_mgn(workbook)
        finally:
            workbook.close()

    def _read_mgn(self, workbook: Workbook) -> None:
        """
        Enriquece los campos utilizando la metadata MGN.

        Esta fuente complementa la información existente,
        sin aplicar reglas de prioridad.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["mgn"],
            "mgn",
        ):
            schema_name = row.get("schema", "")
            table_name = row.get("table", "")
            field_name = row.get("field", "")

            if not schema_name or not table_name or not field_name:
                continue

            field = self._get_field(schema_name, table_name, field_name)

            self._set_if_value(
                field,
                "description",
                row.get("description", ""),
                "description_source",
                "mgn",
            )

            self._set_if_value(
                field,
                "observations",
                row.get("observations", ""),
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
