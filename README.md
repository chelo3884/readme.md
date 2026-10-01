# Residencia – Modelo estructural SAP2000

Proyecto de diseño estructural de una residencia (Quito, Ecuador).
Iteraciones del modelo versionadas con git: cada versión del `.s2k` es un commit.

## Datos de diseño (completar)

| Parámetro | Valor |
|---|---|
| Norma sísmica | NEC __ (indicar edición y capítulo: NEC-SE-DS) |
| Zona / Z | __ |
| Tipo de suelo | __ |
| Sistema estructural / R | __ |
| Importancia I | __ |
| Concreto f'c | __ kgf/cm² |
| Acero fy | __ kgf/cm² |
| Deriva máx. admisible | __ (citar sección de la norma) |

## Unidades

**Tonf, m, kgf/cm²** (indicar en el `.s2k` la unidad base al exportar: `Tonf, m, C`).

## Estructura del repositorio

```
v10/
  modelo.s2k        # modelo exportado desde SAP2000 (File > Export > .s2k)
  tablas/           # tablas exportadas (CSV/Excel): derivas, reacciones, diseño
docs/               # memoria de cálculo, notas, criterios
scripts/
  extraer_tablas_sap.py   # extrae las tablas de SAP2000 a Excel (corre en tu PC)
  calculo_r1.py           # diseño de muros, zapatas, vigas y pedestales (Excel -> calculo_r1.xlsx)
  figuras_r1.py           # figuras de la memoria
  generar_memoria_r1.py   # arma la memoria Rev. 1 a partir de la Rev. 0 y los resultados
```

Cada versión nueva va en su carpeta: `v11/`, `v12/`, ... con la misma estructura.

## Flujo de trabajo

1. En SAP2000: exportar el modelo como `.s2k`, correr análisis/diseño y ejecutar
   `python scripts/extraer_tablas_sap.py --salida vNN/tablas` (requiere `pip install comtypes pandas openpyxl matplotlib pillow`).
2. Subir los archivos a la carpeta de la versión (`vNN/`).
3. En la sesión de Claude Code: se analizan las tablas, se proponen cambios y se genera `v(NN+1)/modelo.s2k`.
4. Importar el nuevo `.s2k` en SAP2000, verificar que abra sin errores, correr análisis y exportar tablas.
5. Repetir.

## Registro de versiones

| Versión | Cambios | Resultado |
|---|---|---|
| v10 | Modelo con cimentación sobre resortes (kv = 4 800 T/m³) | Memoria Rev. 1: `docs/MT-CASA_MR-ESTRUCTURA-R1-01102026.docx` |

## Notas

- No subir `.sdb` ni archivos temporales de SAP2000 (ver `.gitignore`).
- Revisar si el repositorio es privado antes de subir información de clientes.
