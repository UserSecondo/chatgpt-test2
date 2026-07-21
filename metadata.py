"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : metadata.py
Autor    : Nelson David Martínez
Versión  : 5.0

Descripción
-----------
Repositorio central de constantes del proyecto.

Este módulo centraliza toda la configuración estática utilizada por
los diferentes componentes del sistema.

Responsabilidades
-----------------
• Información del proyecto.
• Nombres de archivos.
• Nombres de hojas Excel.
• Modelo lógico del proyecto.
• Mapeo entre cada fuente de datos y el modelo lógico.

IMPORTANTE
----------
Este módulo NO contiene:

- Lectura de archivos.
- Escritura de archivos.
- Reglas de negocio.
- Transformaciones de datos.
===============================================================================
"""

from __future__ import annotations

# =============================================================================
# PROYECTO
# =============================================================================

PROJECT_NAME = "BDD_GEO_DICTIONARY3"

PROJECT_VERSION = "5.0"

PROJECT_AUTHOR = "Nelson David Martínez"

DICTIONARY_TITLE = "Diccionario de Datos"

# =============================================================================
# DIRECTORIOS
# =============================================================================

DIRECTORIES = {

    "input": "input",

    "output": "output",

    "logs": "logs",

    "docs": "docs",

    "tests": "tests"

}

# =============================================================================
# ARCHIVOS DE ENTRADA
# =============================================================================

INPUT_FILENAMES = {

    "oracle":
        "ESQUEMAS_ORACLE_DBGEODIG.xlsx",

    "inventory":
        "inventario_oracle_DBGEODIG_20260708.xlsx",

    "catalog":
        "13052026_CatalogoBDxEsquemas.xlsx",

    "esri":
        "Inventario_Vectorial_DBGEODIG_ESRI_20260707.xlsx",

    "mgn":
        "Diccionario_Datos_MGN_2005.xlsx",

    "template":
        "Plantilla_Diccionario_Datos_DBGEODIG_AJUSTADO.xlsx"

}

# =============================================================================
# FUENTES DE INFORMACIÓN
# =============================================================================

SOURCES = (

    "oracle",

    "inventory",

    "catalog",

    "esri",

    "mgn"

)

# =============================================================================
# HOJAS EXCEL
# =============================================================================

SHEETS = {

    "users":
        "Usuarios",

    "tables":
        "Tablas",

    "oracle_fields":
        "Campos_Oracle",

    "esri_fields":
        "Campos_ESRI",

    "catalog":
        "Esquema_BD_DIG",

    "mgn": [

        "ADMINISTRATIVO",

        "COLOMBIA",

        "MGN",

        "RURAL",

        "URBANO"

    ]

}

# =============================================================================
# MODELO LÓGICO
# =============================================================================

LOGICAL_FIELDS = {

    "schema": "schema",

    "table": "table",

    "field": "field",

    "type": "type",

    "length": "length",

    "precision": "precision",

    "scale": "scale",

    "nullable": "nullable",

    "default": "default",

    "description": "description",

    "responsible": "responsible",

    "domain": "domain",

    "alias": "alias",

    "dataset": "dataset",

    "featureclass": "featureclass",

    "records": "records",

    "created": "created"

}

# =============================================================================
# MAPEO DE COLUMNAS
#
# Cada fuente traduce sus columnas físicas al modelo lógico.
# =============================================================================

COLUMN_MAPS = {

    # -------------------------------------------------------------------------
    # Oracle
    # -------------------------------------------------------------------------

    "users": {

        "USERNAME":
            LOGICAL_FIELDS["schema"],

        "RESPONSABLE":
            LOGICAL_FIELDS["responsible"],

        "REVISION":
            "revision"

    },

    # -------------------------------------------------------------------------

    "tables": {

        "PROPIETARIO":
            LOGICAL_FIELDS["schema"],

        "TABLA":
            LOGICAL_FIELDS["table"],

        "DESCRIPCION":
            LOGICAL_FIELDS["description"]

    },

    # -------------------------------------------------------------------------

    "oracle_fields": {

        "Esquema":
            LOGICAL_FIELDS["schema"],

        "Tabla":
            LOGICAL_FIELDS["table"],

        "Campo":
            LOGICAL_FIELDS["field"],

        "Tipo":
            LOGICAL_FIELDS["type"],

        "Longitud/Precision":
            LOGICAL_FIELDS["length"],

        "Nulable":
            LOGICAL_FIELDS["nullable"],

        "Valor por defecto":
            LOGICAL_FIELDS["default"],

        "Descripcion":
            LOGICAL_FIELDS["description"]

    },

    # -------------------------------------------------------------------------
    # Inventario Oracle
    # -------------------------------------------------------------------------

    "inventory": {

        "PROPIETARIO":
            LOGICAL_FIELDS["schema"],

        "TABLA":
            LOGICAL_FIELDS["table"],

        "REGISTROS":
            LOGICAL_FIELDS["records"],

        "FECHA_CREACION":
            LOGICAL_FIELDS["created"]

    },

    # -------------------------------------------------------------------------
    # Catálogo
    # -------------------------------------------------------------------------

    "catalog": {

        "NOMBRE DEL ESQUEMA":
            LOGICAL_FIELDS["schema"],

        "DESCRIPCION":
            LOGICAL_FIELDS["description"],

        "PERSONA(S) RESPONSABLE ESQUEMA":
            LOGICAL_FIELDS["responsible"]

    },

    # -------------------------------------------------------------------------
    # ESRI
    # -------------------------------------------------------------------------

    "esri_fields": {

        "ESQUEMA":
            LOGICAL_FIELDS["schema"],

        "DATASET":
            LOGICAL_FIELDS["dataset"],

        "FEATURECLASS":
            LOGICAL_FIELDS["featureclass"],

        "Campo":
            LOGICAL_FIELDS["field"],

        "Tipo":
            LOGICAL_FIELDS["type"],

        "Longitud/Precision":
            LOGICAL_FIELDS["length"],

        "Nulable":
            LOGICAL_FIELDS["nullable"],

        "Dominio":
            LOGICAL_FIELDS["domain"],

        "Descripcion":
            LOGICAL_FIELDS["description"]

    },

    # -------------------------------------------------------------------------
    # MGN
    # -------------------------------------------------------------------------

    "mgn": {

        "CAMPO":
            LOGICAL_FIELDS["field"],

        "DESCRIPCIÓN":
            LOGICAL_FIELDS["description"]

    }

}

# =============================================================================
# COLUMNAS OBLIGATORIAS
#
# Utilizadas por excel_utils.open_table() para validar que una hoja contiene
# la información mínima requerida antes de iniciar la lectura.
# =============================================================================

REQUIRED_COLUMNS = {

    "users": (

        "USERNAME",

    ),

    "tables": (

        "PROPIETARIO",
        "TABLA",

    ),

    "oracle_fields": (

        "Esquema",
        "Tabla",
        "Campo",

    ),

    "inventory": (

        "PROPIETARIO",
        "TABLA",

    ),

    "catalog": (

        "NOMBRE DEL ESQUEMA",

    ),

    "esri_fields": (

        "Campo",

    ),

    "mgn": (

        "CAMPO",

    )

}

# =============================================================================
# PRIORIDAD DE LAS FUENTES
#
# Define el orden en que MetadataResolver debe consultar las diferentes
# fuentes para enriquecer la información.
# =============================================================================

PRIORITY = {

    # Responsable del esquema
    "responsible": (

        "catalog",
        "users",

    ),

    # Descripción de tabla
    "table_description": (

        "inventory",
        "catalog",

    ),

    # Descripción de campos
    "field_description": (

        "oracle",
        "esri",
        "mgn",

    )

}

# =============================================================================
# REGLAS DE NEGOCIO
#
# Estas constantes únicamente documentan las reglas.
# La implementación pertenece a metadata_resolver.py
# =============================================================================

BUSINESS_RULES = {

    "RN-001":
        "Solo documentar esquemas marcados para revisión.",

    "RN-002":
        "Solo documentar tablas pertenecientes a los esquemas seleccionados.",

    "RN-003":
        "Los archivos complementarios únicamente enriquecen información.",

    "RN-004":
        "El responsable del esquema se obtiene por prioridad de fuentes.",

    "RN-005":
        "La descripción del campo se obtiene por prioridad Oracle → ESRI → MGN.",

    "RN-006":
        "La descripción de la tabla se obtiene por prioridad Inventario → Catálogo.",

    "RN-007":
        "Si una fuente no tiene información continuar con la siguiente."

}

# =============================================================================
# VALORES CONSTANTES
# =============================================================================

VALUES = {

    "yes": "SI",

    "no": "NO",

    "true": True,

    "false": False,

    "empty": "",

    "none": None,

    "unknown": "No disponible"

}

# =============================================================================
# MENSAJES
#
# Utilizados por logger.py y main.py
# =============================================================================

MESSAGES = {

    "project_start":
        "Iniciando generación del diccionario de datos.",

    "loading_configuration":
        "Cargando configuración del proyecto...",

    "reading_sources":
        "Leyendo archivos de entrada...",

    "resolving_metadata":
        "Resolviendo metadatos...",

    "writing_dictionary":
        "Generando diccionarios de datos...",

    "finished":
        "Proceso finalizado correctamente.",

    "error":
        "Se produjo un error durante la ejecución."

}

# =============================================================================
# FORMATOS DE FECHA
# =============================================================================

DATE_FORMATS = {

    "report":

        "%Y%m%d_%H%M%S",

    "log":

        "%Y-%m-%d %H:%M:%S"

}

# =============================================================================
# EXTENSIONES DE ARCHIVO
# =============================================================================

FILE_EXTENSIONS = {

    "excel": ".xlsx",

    "log": ".log"

}

# =============================================================================
# NOMBRES DE ARCHIVOS DE SALIDA
# =============================================================================

OUTPUT_FILENAMES = {

    "dictionary":

        "Diccionario_Datos_DBGEODIG.xlsx",

    "log":

        "BDD_GEO_DICTIONARY.log"

}

# =============================================================================
# TIPOS DE DATO ORACLE
#
# Normalización opcional utilizada por MetadataResolver.
# =============================================================================

ORACLE_TYPES = {

    "VARCHAR2": "VARCHAR2",

    "CHAR": "CHAR",

    "NUMBER": "NUMBER",

    "DATE": "DATE",

    "CLOB": "CLOB",

    "BLOB": "BLOB",

    "FLOAT": "FLOAT",

    "TIMESTAMP": "TIMESTAMP"

}

# =============================================================================
# ESTADOS
# =============================================================================

STATUS = {

    "pending": "PENDIENTE",

    "processing": "PROCESANDO",

    "completed": "COMPLETADO",

    "error": "ERROR"

}

# =============================================================================
# FIN DEL ARCHIVO
# =============================================================================