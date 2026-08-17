# Práctica: Lenguaje M — Explorando el Editor Avanzado de Power Query

## 1. Dataset utilizado

Para realizar esta práctica utilicé el dataset **Locales Bailables**, publicado por **Buenos Aires Data — Gobierno de la Ciudad Autónoma de Buenos Aires**.

**Fuente:** Buenos Aires Data
**Organismo responsable:** Ministerio de Justicia | Agencia Gubernamental de Control (AGC) | Dirección General Habilitaciones y Permisos
**Formato utilizado:** XLSX
**Dataset:** Locales Bailables

El dataset contiene información sobre locales inscriptos y reinscriptos en el Registro Público de Lugares Bailables y Eventos Masivos de la Ciudad de Buenos Aires.

## 2. Justificación de la elección del dataset

Elegí este dataset porque proviene de una **fuente pública y oficial** y presenta características adecuadas para practicar limpieza y transformación de datos en Power Query.

La base contiene múltiples columnas con diferentes tipos de información, como nombres de locales, razón social, domicilio, capacidad, estado y fechas de inscripción.

Además, presenta situaciones que requieren transformaciones antes de realizar un análisis. Por ejemplo, la columna `Capacidad` contenía valores de texto como `428 personas`, por lo que fue necesario limpiar el texto y convertir posteriormente la columna a un tipo de dato numérico.

También contiene valores nulos y nombres de columnas que pueden mejorarse para facilitar su interpretación.

Por estos motivos considero que el dataset representa un caso realista para practicar procesos de preparación y transformación de datos.

## 3. Transformaciones realizadas

Se realizaron las siguientes transformaciones utilizando Power Query:

### Transformación 1 — Limpieza y cambio de tipo de dato

La columna `Capacidad` contenía valores como:

`428 personas`

`901 personas`

`1372 personas`

Primero se eliminó el texto ` personas` mediante la función de reemplazo de valores.

Luego se cambió el tipo de dato de la columna `Capacidad` a **Número entero**.

### Transformación 2 — Filtrado de filas

Se aplicó un filtro sobre la columna `Estado` para conservar únicamente los registros correspondientes a:

`RENOVADO 2026`

De esta manera se trabajó únicamente con los locales que presentan dicho estado.

### Transformación 3 — Renombrado de columna

La columna:

`Nº REG`

fue renombrada como:

`ID_Registro`

El objetivo fue utilizar un nombre más descriptivo y sencillo para identificar el campo.

## 4. Código M generado

Luego de realizar las transformaciones mediante la interfaz visual de Power Query, se accedió al **Editor Avanzado** para analizar y modificar manualmente el código M.

El código final utilizado fue:

```powerquery
let
    Origen = Excel.Workbook(File.Contents("C:\Documentos\Gerardo\Business Analytic\Modulo_6\Ejercicio_2\locales_bailables.xlsx"), null, true),
    #"Lista de locales inscriptos!_xlnm.Print_Area_DefinedName" = Origen{[Item="Lista de locales inscriptos!_xlnm.Print_Area",Kind="DefinedName"]}[Data],
    #"Encabezados promovidos" = Table.PromoteHeaders(#"Lista de locales inscriptos!_xlnm.Print_Area_DefinedName", [PromoteAllScalars=true]),
    #"Valor reemplazado" = Table.ReplaceValue(#"Encabezados promovidos"," personas","",Replacer.ReplaceText,{"Capacidad"}),
    #"Tipo cambiado" = Table.TransformColumnTypes(#"Valor reemplazado",{{"Capacidad", Int64.Type}}),
    #"Filas filtradas" = Table.SelectRows(#"Tipo cambiado", each ([Estado] = "RENOVADO 2026")),

    // Renombro la columna de registro para darle un nombre más claro y descriptivo
    #"MiTransformacionManual" = Table.RenameColumns(#"Filas filtradas",{{"Nº             REG", "ID_Registro"}})
in
    #"MiTransformacionManual"
```

## 5. Modificación manual realizada

Desde el Editor Avanzado se modificó manualmente el nombre del último paso.

El nombre generado originalmente por Power Query era:

```powerquery
#"Columnas con nombre cambiado"
```

y fue reemplazado por:

```powerquery
#"MiTransformacionManual"
```

También se actualizó la referencia correspondiente después de `in`:

```powerquery
in
    #"MiTransformacionManual"
```

Finalmente, se agregó el siguiente comentario para documentar el objetivo de la transformación:

```powerquery
// Renombro la columna de registro para darle un nombre más claro y descriptivo
```

Después de realizar estas modificaciones, el código se ejecutó correctamente en Power Query.

## 6. ¿Por qué es útil para un analista de datos entender la estructura `let … in`?

Entender la estructura `let … in` permite comprender cómo Power Query ejecuta las transformaciones de forma secuencial.

Dentro de `let` se definen los diferentes pasos de transformación. Cada paso puede utilizar como entrada el resultado del paso anterior, generando una cadena de transformaciones.

Por ejemplo:

`Origen → Encabezados promovidos → Valor reemplazado → Tipo cambiado → Filas filtradas → MiTransformacionManual`

La expresión `in` determina cuál de esos pasos será devuelto como **resultado final de la consulta**.

Comprender esta estructura permite al analista modificar consultas manualmente, identificar errores, agregar transformaciones más complejas y entender qué está haciendo Power Query detrás de la interfaz visual.

## 7. ¿Qué significa que el Lenguaje M sea Case Sensitive?

Que el Lenguaje M sea **Case Sensitive** significa que distingue entre letras mayúsculas y minúsculas.

Por ejemplo:

```powerquery
Table.SelectRows
```

no es lo mismo que:

```powerquery
table.selectrows
```

La primera expresión corresponde a una función válida de Power Query, mientras que la segunda puede generar un error porque el nombre de la función no fue escrito respetando correctamente las mayúsculas y minúsculas.

La consecuencia práctica es que al escribir o modificar código manualmente en el Editor Avanzado es necesario respetar exactamente los nombres de funciones, variables y pasos.

## 8. Conclusión

Esta práctica permitió observar que las transformaciones realizadas desde la interfaz gráfica de Power Query generan automáticamente código en Lenguaje M.

A través del Editor Avanzado fue posible identificar cada uno de los pasos, comprender la estructura `let … in` y modificar manualmente una transformación sin alterar el funcionamiento de la consulta.

Esto demuestra la importancia de conocer Lenguaje M para poder desarrollar procesos de transformación más flexibles, reutilizables y fáciles de depurar.
