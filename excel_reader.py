"""
===============================================================================
Proyecto : BDD_GEO_DICTIONARY3
Archivo  : excel_reader.py
Versión  : 6.0

Punto de entrada del lector de Excel.

Este módulo únicamente compone las tres piezas en las que se dividió
la lectura de Excel:

- ExcelReaderBase    : utilidades comunes, índices internos y
                       asignación de valores según prioridad de fuentes.
- ExcelReaderOracle   : lectura de Oracle, inventario y catálogo.
- ExcelReaderSources  : lectura de ESRI, MGN y validación del modelo.

No contiene lógica propia; toda la implementación vive en los
módulos anteriores.
===============================================================================
"""

from __future__ import annotations

from excel_reader_base import ExcelReaderBase
from excel_reader_oracle import ExcelReaderOracle
from excel_reader_sources import ExcelReaderSources


class ExcelReader(ExcelReaderBase, ExcelReaderOracle, ExcelReaderSources):
    """
    Construye el modelo ProjectMetadata leyendo todas
    las fuentes Excel del proyecto.
    """

    pass
