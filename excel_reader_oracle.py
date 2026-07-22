"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : excel_reader_oracle.py
Versión  : 5.1

Lectura de las fuentes Oracle, Inventario y Catálogo.

Responsabilidades
-----------------
- Leer usuarios Oracle.
- Leer tablas Oracle.
- Leer campos Oracle.
- Leer inventario Oracle.
- Leer catálogo de esquemas.

No realiza:
- Resolución de conflictos entre fuentes.
- Aplicación de reglas de negocio.
- Escritura del diccionario de datos.
===============================================================================
"""

from __future__ import annotations

from openpyxl.workbook.workbook import Workbook

from metadata import SHEETS


class ExcelReaderOracle:
    """
    Implementa la lectura de Oracle, inventario y catálogo.
    """

    ###########################################################################
    # Oracle
    ###########################################################################

    def _load_oracle(self) -> None:
        """
        Procesa el archivo de metadatos Oracle.
        """

        workbook = self._open_workbook("oracle")

        try:
            self._read_users(workbook)
            self._read_tables(workbook)
            self._read_oracle_fields(workbook)
        finally:
            workbook.close()

    def _read_users(self, workbook: Workbook) -> None:
        """
        Lee la hoja de usuarios Oracle.

        Crea los esquemas e incorpora el responsable
        cuando esté disponible.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["users"],
            "users",
        ):
            schema_name = row.get("schema", "")

            if not schema_name:
                continue

            documented = row.get("revision", "") == "SI"

            self._documented_schemas[schema_name] = documented

            if not documented:
                continue

            schema = self._get_schema(schema_name)

            responsible = row.get("responsible", "")

            self._assign_by_priority(
                schema,
                "responsible",
                "responsible_source",
                "responsible",
                responsible,
                "users",
            )

    ###########################################################################
    # Tablas Oracle
    ###########################################################################

    def _read_tables(self, workbook: Workbook) -> None:
        """
        Lee la hoja de tablas Oracle.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["tables"],
            "tables",
        ):
            schema_name = row.get("schema", "")
            table_name = row.get("table", "")

            if not schema_name or not table_name:
                continue

            if not self._is_documented(schema_name):
                continue

            table = self._get_table(schema_name, table_name)

            description = row.get("description", "")

            self._assign_by_priority(
                table,
                "description",
                "description_source",
                "table_description",
                description,
                "oracle",
            )

    ###########################################################################
    # Campos Oracle
    ###########################################################################

    def _read_oracle_fields(self, workbook: Workbook) -> None:
        """
        Lee los campos Oracle.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["oracle_fields"],
            "oracle_fields",
        ):
            schema_name = row.get("schema", "")
            table_name = row.get("table", "")
            field_name = row.get("field", "")

            if not schema_name or not table_name or not field_name:
                continue

            if not self._is_documented(schema_name):
                continue

            field = self._get_field(schema_name, table_name, field_name)

            value = row.get("type", "")
            if value:
                field.data_type = value
                field.data_type_source = "oracle"
                field.oracle_type = value

            value = row.get("length", "")
            if value:
                field.length = value
                field.oracle_length = str(value)

            value = row.get("nullable", "")
            if value:
                field.nullable = value

            value = row.get("description", "")

            if not self._is_placeholder_description(value, field_name):
                self._assign_by_priority(
                    field,
                    "description",
                    "description_source",
                    "field_description",
                    value,
                    "oracle",
                )

    ###########################################################################
    # Utilidad de asignación
    ###########################################################################

    @staticmethod
    def _set_if_value(
        obj: object,
        attribute: str,
        value: str,
        source_attribute: str | None = None,
        source: str | None = None,
    ) -> None:
        """
        Asigna un valor únicamente cuando éste no es vacío.

        Si se indica un atributo de origen, también registra
        la fuente del dato.
        """

        if value is None:
            return

        value = value.strip()

        if not value:
            return

        setattr(obj, attribute, value)

        if source_attribute and source:
            setattr(obj, source_attribute, source)

    ###########################################################################
    # Inventario Oracle
    ###########################################################################

    def _load_inventory(self) -> None:
        """
        Procesa el inventario Oracle.
        """

        workbook = self._open_workbook("inventory")

        try:
            self._read_inventory(workbook)
        finally:
            workbook.close()

    def _read_inventory(self, workbook: Workbook) -> None:
        """
        Lee la información del inventario Oracle.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["inventory"],
            "inventory",
        ):
            schema_name = row.get("schema", "")
            table_name = row.get("table", "")

            if not schema_name or not table_name:
                continue

            if not self._is_documented(schema_name):
                continue

            table = self._get_table(schema_name, table_name)

            self._set_if_value(
                table,
                "description",
                row.get("description", ""),
                "description_source",
                "inventory",
            )

    ###########################################################################
    # Catálogo
    ###########################################################################

    def _load_catalog(self) -> None:
        """
        Procesa el catálogo de esquemas.
        """

        workbook = self._open_workbook("catalog")

        try:
            self._read_catalog(workbook)
        finally:
            workbook.close()

    def _read_catalog(self, workbook: Workbook) -> None:
        """
        Lee el catálogo de esquemas.
        """

        for row in self._read_sheet(
            workbook,
            SHEETS["catalog"],
            "catalog",
        ):
            schema_name = row.get("schema", "")

            if not schema_name:
                continue

            if not self._is_documented(schema_name):
                continue

            schema = self._get_schema(schema_name)

            self._assign_by_priority(
                schema,
                "responsible",
                "responsible_source",
                "responsible",
                row.get("responsible", ""),
                "catalog",
            )

            self._set_if_value(
                schema,
                "description",
                row.get("description", ""),
                "description_source",
                "catalog",
            )
