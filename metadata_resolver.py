"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : metadata_resolver.py
Versión  : 6.0

Cierra la resolución de metadatos a partir del modelo leído.

La prioridad entre fuentes (RN-004, RN-005 y RN-006) ya se aplica
durante la lectura, en excel_reader_oracle.py y excel_reader_sources.py,
mediante ExcelReaderBase._assign_by_priority. Este módulo se encarga
del paso final:

- Detectar los elementos que ninguna fuente logró completar.
- Rellenarlos con el valor "No disponible" para que el diccionario
  de datos generado no tenga celdas vacías.
- Registrar una advertencia por cada elemento incompleto.

No realiza:
- Lectura de archivos Excel.
- Escritura del diccionario de datos.
- Comparación entre fuentes (eso ya ocurrió durante la lectura).
===============================================================================
"""

from __future__ import annotations

from metadata import VALUES
from models import ProjectMetadata

_UNKNOWN = VALUES["unknown"]


class MetadataResolver:
    """
    Completa los valores que ninguna fuente pudo resolver.
    """

    def resolve(self, project: ProjectMetadata) -> ProjectMetadata:
        """
        Resuelve los metadatos del proyecto sin crear un nuevo modelo.
        """

        self._resolve_schema_responsibles(project)
        self._resolve_table_descriptions(project)
        self._resolve_field_descriptions(project)

        return project

    def _resolve_schema_responsibles(self, project: ProjectMetadata) -> None:
        """
        Completa el responsable de los esquemas que quedaron sin
        información en ninguna fuente (catálogo ni usuarios Oracle).
        """

        for schema in project.schemas:

            if schema.responsible:
                continue

            schema.responsible = _UNKNOWN

            project.warnings.append(
                f"El esquema '{schema.schema_name}' no tiene responsable "
                "en ninguna fuente (catálogo ni usuarios Oracle)."
            )

    def _resolve_table_descriptions(self, project: ProjectMetadata) -> None:
        """
        Completa la descripción de las tablas que quedaron sin
        información en ninguna fuente (Oracle ni inventario).
        """

        for schema in project.schemas:
            for table in schema.tables:

                if table.description:
                    continue

                table.description = _UNKNOWN

                project.warnings.append(
                    f"La tabla '{table.full_name}' no tiene descripción "
                    "en ninguna fuente (Oracle ni inventario)."
                )

    def _resolve_field_descriptions(self, project: ProjectMetadata) -> None:
        """
        Completa la descripción de los campos que quedaron sin
        información en ninguna fuente (Oracle, ESRI ni MGN).
        """

        for schema in project.schemas:
            for table in schema.tables:
                for field in table.fields:

                    if field.description:
                        continue

                    field.description = _UNKNOWN

                    project.warnings.append(
                        f"{schema.schema_name}.{table.table_name}."
                        f"{field.field_name} no tiene descripción en "
                        "ninguna fuente (Oracle, ESRI ni MGN)."
                    )
