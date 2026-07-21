"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : models.py
Autor    : Nelson David Martínez
Versión  : 5.0

Descripción
-----------
Modelo de dominio del proyecto.

Este módulo define las entidades utilizadas durante toda la ejecución.

IMPORTANTE
-----------
No contiene:

- Lectura de archivos Excel.
- Reglas de negocio.
- Escritura de archivos.
- Dependencias con Oracle, ESRI o MGN.
===============================================================================
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import List


# =============================================================================
# CAMPO
# =============================================================================

@dataclass
class FieldInfo:
    """
    Representa un campo perteneciente a una tabla.
    """

    # Información básica

    field_name: str

    # Tipo de dato

    data_type: str = ""

    data_type_source: str = ""

    # Longitud

    length: str = ""

    # Nulabilidad

    nullable: str = ""

    # Dominio

    domain: str = ""

    # Alias

    alias: str = ""

    # Descripción

    description: str = ""

    description_source: str = ""

    # Observaciones

    observations: str = ""


# =============================================================================
# TABLA
# =============================================================================

@dataclass
class TableInfo:
    """
    Representa una tabla perteneciente a un esquema.
    """

    schema_name: str

    table_name: str

    description: str = ""

    description_source: str = ""

    responsible: str = ""

    responsible_source: str = ""

    fields: List[FieldInfo] = field(default_factory=list)


# =============================================================================
# ESQUEMA
# =============================================================================

@dataclass
class SchemaInfo:
    """
    Representa un esquema Oracle.
    """

    schema_name: str

    description: str = ""

    description_source: str = ""

    responsible: str = ""

    responsible_source: str = ""

    tables: List[TableInfo] = field(default_factory=list)


# =============================================================================
# RESUMEN
# =============================================================================

@dataclass
class GenerationSummary:
    """
    Resumen de la ejecución.
    """

    processed_schemas: int = 0

    processed_tables: int = 0

    processed_fields: int = 0

    generated_files: int = 0

    warnings: int = 0

    errors: int = 0

    start_time: datetime | None = None

    end_time: datetime | None = None

    execution_time: timedelta = timedelta()


# =============================================================================
# PROYECTO
# =============================================================================

@dataclass
class ProjectMetadata:
    """
    Contenedor principal del proyecto.
    """

    schemas: List[SchemaInfo] = field(default_factory=list)

    summary: GenerationSummary = field(
        default_factory=GenerationSummary
    )

    warnings: List[str] = field(default_factory=list)

    errors: List[str] = field(default_factory=list)