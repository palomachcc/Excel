# Excel ![Badge en Desarollo](https://img.shields.io/badge/STATUS-EN%20DESAROLLO-green)

## :hammer:Repaso de algunas funciones utiles de excel.

- `Parte 1`: Cleaning and manipulating text. Concatenar o unir textos, extraer partes de palabras. Removing and replacing text characters
- `Parte 2`: Format data. Varias funciones utiles para cuando laburas con fechas, calculos con dias habiles o vacaciones etc.
- `Parte 3`: descripción de la funcionalidade 2a relacionada con la funcionalidad 2
- `Parte 4`: descripción de la funcionalidad 3

-------
A la hora de trabajar con datos es necesario que estén limpios y ordenados para asegurarnos de que sean coherentes. Esto también ayuda a reducir la posibilidad de errores y mejora la eficiencia a la hora de analizarlos. De esta manera es mas fácil detectar patrones o tendencias y poder así tomar decisiones informadas.

### Algunos ejercicios como ejemplo:

#### Ejemplo Invoice Report
El objetivo de esta tarea es usar las herramientas de excel para limpiar y ordenar los datos antes de empezar a trabajar con la información.

Contexto: Recibo un archivo (Hoja Supplier Invoice Statement) de un proveedor con los pagos realizados en el mes. La idea es ingresar esos datos a mi sistema. Todo se descarga y se realiza en una hoja de excel. 

[Invoice Report Inicial.xlsx](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/4d8045d9-b902-44c5-af4f-fb069d7a695d/Invoice_Report_Inicial.xlsx)

Muchos de los datos que necesito están agrupados en una misma columna o de a partes y también en formatos diferentes a como los tengo en el sistema. Otras celdas contienen caracteres raros sin ningún uso. 

Excel tiene una variedad de funciones, que son simplemente procesos predefinidos por el programa.

![image](https://user-images.githubusercontent.com/110131341/226074495-872e0e63-5b2e-4580-9f83-7cb564cccc41.png)

Para empezar con la tarea, con las siguientes funciones se resuelven los datos de texto de manera rápida:

- Unir de texto (CONCATENATE, &, CONCAT, TEXTJOIN)
- Separar texto (LEFT, RIGHT, MID, FIND,LEN)
- info para combinar con otras funciones(FIND,LEN)
- Limpieza y orden (CLEAN,TRIM UPPER,LOWER,PROPER)
- Remover y reemplazar (SUBSTITUTE)

[Invoice Report (2).xlsx](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/a75e61ad-43e2-4f7e-a5fb-7faf31a237e3/Invoice_Report_(2).xlsx)

Lo siguiente que queda por hacer es convertir tipos de datos, generar fechas validas, y hacer cálculos con fechas (teniendo en cuenta días hábiles ) . Lo realizamos con las siguientes funciones:

- Conversion tipo (VALUE, TEXT)
- Tiempo y fechas (DAY,YEAR,MONTH,NOW,TODAY, DATE)
- Cálculos con fechas (DAYS, NETWORKDAYS, WORKDAY)

[Invoice Report (3).xlsx](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/161d3b72-1d30-400f-8ddb-03a1b3d95c77/Invoice_Report_(3).xlsx)
