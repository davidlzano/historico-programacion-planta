# Histórico versionado de programación de planta

Respaldo diario de un archivo de programación editado a mano, y consolidación de esos respaldos en una base única con línea de tiempo consultable.

Este repositorio es una **reimplementación demostrativa** de un sistema que puse en producción sobre la programación de 22 centros de trabajo. El código aquí publicado es original, trabaja con datos sintéticos y no contiene información de la empresa.

---

## El problema

La programación de producción vivía en un solo archivo de Excel, editado a diario por varias personas. Ese archivo tenía dos características incómodas:

**No había historia.** Solo existía el estado de hoy. Preguntas como "¿qué habíamos programado para esta OP la semana pasada?" o "¿cuántas veces se ha reprogramado esta máquina?" no tenían respuesta, porque el dato anterior se sobrescribía.

**No había red de seguridad.** Un archivo compartido que edita mucha gente termina rompiéndose.

Eso segundo dejó de ser hipotético. El 31 de agosto alguien eliminó una columna y dejó 1.850 celdas con referencias rotas, destruyendo la secuencia de programación de las dos máquinas principales. Se pudo reconstruir porque existía el respaldo de dos días antes.

## El enfoque

Dos piezas que se complementan.

### 1. Respaldo diario

Una tarea programada copia el archivo cada mañana a una carpeta de histórico, organizada automáticamente por mes y año. Reintenta si el archivo está abierto en ese momento, y registra en el log el contexto de ejecución completo.

Ese detalle del contexto no es decorativo: cuando una tarea programada falla en un servidor, saber con qué usuario e intérprete corrió suele ser la diferencia entre resolverlo en minutos o pasar la tarde adivinando.

### 2. Consolidación con versionado

Los respaldos por sí solos son una carpeta con cientos de archivos. La segunda pieza los convierte en una base consultable con dos capas:

| Hoja | Contenido |
|---|---|
| `VIGENTE` | El estado actual de cada registro. La tabla para consultar y relacionar. |
| `HISTORICO` | Una fila cada vez que un registro nace o cambia. |
| `AVISOS` | Hojas que no se pudieron leer. La que hay que revisar. |

---

## Las tres decisiones que sostienen el sistema

### Mapeo por nombre de encabezado, no por posición

El archivo origen lo edita a mano el área de producción. Las columnas se mueven, se renombran y aparecen nuevas. Un mismo campo aparece como `TIROS`, `CANTIDAD` o `CANTIDAD PEDIDA` según quién editó la hoja, y el encabezado no siempre está en la fila 1.

Leer por posición funciona hasta el primer día en que alguien inserta una columna. A partir de ahí carga datos en el campo equivocado **sin lanzar ningún error**, que es la peor forma de fallar: nadie se entera hasta que alguien nota que las cifras no cuadran, semanas después.

El consolidador mantiene una lista de sinónimos por campo, busca la fila de encabezados en las primeras doce, e ignora los encabezados que no reconoce en lugar de romper la carga.

### Histórico por cambios, no por fotos completas

Guardar la foto entera de cada día es simple de leer y satura el límite de filas de Excel en cerca de un año. Guardar solo cuando algo nace o cambia conserva exactamente la misma información: para saber cómo estaba un registro en cualquier fecha, se toma su última versión con snapshot menor o igual a esa fecha.

Lo que decide qué cuenta como "cambio" son los campos de negocio. Los de trazabilidad —fecha del snapshot, archivo de origen— quedan fuera a propósito. Si estuvieran incluidos, cada día generaría una versión nueva de absolutamente todo y el modo dejaría de servir.

En este repositorio, con veinte días simulados, la reducción ronda el **34%**. Sobre la programación real, que cambia bastante menos entre días, fue del **82%** conservando el mismo histórico.

### Una hoja ilegible no es una fila borrada

Esta es la parte fina.

Al reconstruir el estado vigente, el consolidador avanza día por día. Si una fila deja de aparecer, lo normal es concluir que la borraron y sacarla del vigente. Pero hay otra explicación: que esa hoja no se haya podido leer ese día.

Confundir ambos casos tiene una consecuencia grave. Un archivo con una hoja corrupta un martes borraría del estado vigente toda la programación de esa máquina, y el sistema no daría ninguna señal de que algo pasó.

Por eso el consolidador registra qué hojas leyó efectivamente cada día, y **solo purga llaves dentro de las hojas que sí pudo leer**. Las que faltaron se dejan intactas y quedan reportadas en `AVISOS`.

---

## Ejecución

Requiere Python 3.10 o superior.

```bash
pip install -r requirements.txt

python generar_snapshots.py   # crea 20 días de respaldos con layout variable
python consolidador.py        # produce BD_PROGRAMACION.xlsx
```

El generador reproduce a propósito el problema real: cada archivo tiene las columnas en otro orden, con nombres distintos, encabezados en filas diferentes, y alguna hoja ocasionalmente ilegible. Abre un par de archivos de `snapshots/` y compáralos para verlo.

Salida típica:

```
Procesando 14 respaldos...

Filas leidas                  323
Historico (CAMBIOS)           212  (34% menos)
Registros vigentes             24  (3 son eventos, no produccion)
Avisos de lectura               3

Hojas que no se pudieron leer:
  PROGRAMACION 07092026.xlsx / CONV.: Sin encabezados reconocibles
```

Las hojas ilegibles aparecen reportadas y sus registros **no** se dan por eliminados.

---

## Estructura

| Archivo | Contenido |
|---|---|
| `generar_snapshots.py` | Crea los respaldos de ejemplo con layout variable |
| `consolidador.py` | Mapeo por encabezado, versionado y reconstrucción del vigente |

---

## Diferencias con la versión en producción

| | Aquí | Producción |
|---|---|---|
| Origen | archivos generados | `.xlsm` en carpeta de red |
| Centros de trabajo | 5 | 22 |
| Respaldo | fuera de alcance | tarea programada diaria con reintentos |
| Sinónimos por campo | 10 campos | 26 campos |
| Casos especiales | ninguno | mapeo forzado para hojas con encabezados desplazados |

Sobre ese último punto: en el archivo real hay hojas donde los títulos están corridos una columna respecto a los datos, y otras donde los encabezados de turno viven en una fila distinta al resto. Esos casos no se resuelven con sinónimos y llevan un mapeo explícito por hoja.

---

## Nota sobre la construcción

El consolidador se desarrolló con asistencia de IA. El diseño del sistema —guardar respaldos diarios antes de necesitarlos, mapear por encabezado en vez de por posición, y no confundir una hoja ilegible con una fila borrada— responde a problemas concretos observados en la operación, y las cifras reportadas fueron verificadas contra los archivos reales.

---

## Posibles extensiones

- Detección de solapamientos de tiempo entre tramos de una misma máquina
- Alerta cuando un registro cambia más de N veces en una semana, como señal de reprogramación excesiva
- Salida a base de datos en vez de Excel, para eliminar el límite de filas

---

## Licencia

MIT
