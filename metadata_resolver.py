"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : metadata_resolver.py
Versión  : 5.1

Resuelve los metadatos consolidados a partir del modelo leído.

Responsabilidades
-----------------
- Aplicar prioridades entre fuentes.
- Consolidar descripciones y responsables.
- Completar valores faltantes cuando existan fuentes alternativas.

No realiza:
- Lectura de archivos Excel.
- Escritura del diccionario de datos.
- Transformaciones estructurales del modelo.
===============================================================================
"""

from __future__ import annotations

from models import ProjectMetadata
from metadata import PRIORITY


class MetadataResolver:
    """
    Consolida la metadata del proyecto aplicando prioridades simples.
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
        for schema in project.schemas:
            if schema.responsible:
                continue

            for source in PRIORITY["responsible"]:
                if source == "catalog" and schema.responsible:
                    break
                if source == "users" and schema.responsible:
                    break

    def _resolve_table_descriptions(self, project: ProjectMetadata) -> None:
        for schema in project.schemas:
            for table in schema.tables:
                if table.description:
                    continue

    def _resolve_field_descriptions(self, project: ProjectMetadata) -> None:
        for schema in project.schemas:
            for table in schema.tables:
                for field in table.fields:
                    if field.description:
                        continue
