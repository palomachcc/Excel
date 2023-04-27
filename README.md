# Excel ![Badge en Desarollo](https://img.shields.io/badge/STATUS-EN%20DESAROLLO-green)

## :hammer:Repaso de algunas herramientas básicas de excel para limpieza y orden de datos.

- `Parte 1`: Cleaning and manipulating text. Concatenar o unir textos, extraccion de partes. Removing and replacing text characters
- `Parte 2`: Format data. Varias funciones utiles para cuando laburas con fechas, calculos con dias habiles o vacaciones etc.
- `Parte 3`: Cell referencing and naming; Creating named ranges; Managing named ranges; Calculations with named ranges; and Automating processes with named ranges.
- `Parte 4`: Tablas, referencias estructuradas.
- `Parte 5`: Operaciones logicas con IF, AND, OR. VLOOKUP, XLOOKUP, INDEX, MATCH.
-------
A la hora de trabajar con datos es necesario que estén limpios y ordenados para asegurarnos de que sean coherentes. Esto también ayuda a reducir la posibilidad de errores y mejora la eficiencia a la hora de analizarlos. De esta manera es mas fácil detectar patrones o tendencias y poder así tomar decisiones informadas.

### Ejemplo de un reporte de pagos.

Contexto: Recibo un archivo (Hoja Supplier Invoice Statement) de un proveedor con los pagos realizados en el mes. La idea es ingresar esos datos a mi sistema. Todo se descarga y se realiza en una hoja de excel. 
El objetivo de esta tarea es usar las herramientas de excel para limpiar y ordenar los datos antes de empezar a trabajar con la información.

[Invoice Report Inicial.xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%201/Invoice%20Report%20Inicial.xlsx)

Muchos de los datos que necesito están agrupados en una misma columna o de a partes y también en formatos diferentes a como los tengo en el sistema. Otras celdas contienen caracteres raros sin ningún uso. 



#### Funciones
Excel tiene una variedad de funciones, que son simplemente procesos predefinidos por el programa.

![image](https://user-images.githubusercontent.com/110131341/226074495-872e0e63-5b2e-4580-9f83-7cb564cccc41.png)

Para empezar con la tarea, con las siguientes funciones se resuelven los datos de texto de manera rápida:

- Union de texto (CONCATENATE, &, CONCAT, TEXTJOIN)
- Separar texto (LEFT, RIGHT, MID, FIND,LEN)
- info para combinar con otras funciones(FIND,LEN)
- Limpieza y orden (CLEAN,TRIM UPPER,LOWER,PROPER)
- Remover y reemplazar (SUBSTITUTE)

[Invoice Report (2).xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%201/Invoice%20Report%20(2).xlsx)

Lo siguiente que queda por hacer es convertir tipos de datos, generar fechas validas, y hacer cálculos con fechas (teniendo en cuenta días hábiles ). Lo realizamos con las siguientes funciones:

- Conversion tipo (VALUE, TEXT)
- Tiempo y fechas (DAY,YEAR,MONTH,NOW,TODAY, DATE)
- Cálculos con fechas (DAYS, NETWORKDAYS, WORKDAY)

[Invoice Report (3).xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%202/Invoice%20Report%20(3).xlsx)

Una vez transformados y reordenados los datos, se puede pasar a los cálculos.



#### Rangos de datos
Excel tiene la opción de nombrar celdas, rangos de celdas, tablas, formulas y constantes. 

<img src="https://user-images.githubusercontent.com/110131341/226122239-78d5d04d-e242-4fb3-95b9-f0794540ece8.png " width=40% height=40%>
Esto facilita la lectura y manejo de formulas. 

Por cada día de retraso en el pago se cobra una penalidad (Penalty Rate) y una taza fija de $2 por día (Flat Rate). Tenemos que calcular si hubo retraso y, en tal caso, el recargo correspondiente (Late Charge). 

Por otro lado, nos falta completar la hoja “Recon Analysis”. Tenemos que verificar si el pago total del proveedor coincide con el de nuestro sistema y calcular cuantas facturas corresponden a Sydney y cuantas a Melbourne.

Para esto usamos:

- Condicionales (IF, COUNTIF, SUMIF)
- Funciones fecha
- Nombre de celdas
- Nombre de constantes (flat_rate)
- Nombres definidos (Named Ranges). Administrador de nombres, cálculos y automatizado de acciones con rangos.

[Invoice Report (4).xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%203/Invoice%20Report%20(4).xlsx)

Hasta aca, lo básico. El plan es que el trabajo se actualice a medida que se agregan datos mensualmente. Con lo visto hasta ahora, eso debe realizarse manualmente.



#### Tablas

Para facilitar la administración, extracción de información y el análisis de un grupo de datos relacionados es conveniente tener un proceso estructurado y automatizado. Así se evitan errores y se ahorra tiempo. Podemos usar Macros o simplemente crear Tablas. En este caso voy a usar Tablas. 

Los datos que se agregan a una Tabla toman su misma estructura, formato condicional y formulas, entre otras cosas. De este modo, también, se actualizan los rangos.

![image1](https://user-images.githubusercontent.com/110131341/226460271-3fe22d0f-2229-4930-81e3-25e25be01f26.png)

Una vez creada la tabla podemos ordenar y filtrar los datos según lo que necesitemos.

![image2](https://user-images.githubusercontent.com/110131341/226460512-8cab21e6-d66a-4dcd-8b11-7da63f737061.png)

Tambien podemos agregar slicers, con la opción de Segmentación de datos, que permiten insertar filtros de una manera interactiva. Además del filtrado rápido, las segmentaciones de datos también indican el estado de filtrado actual, lo que facilita la comprensión de lo que se muestra exactamente en ese momento.

![image3](https://user-images.githubusercontent.com/110131341/226460687-01678920-5773-4af3-b91c-35d0587440dd.png)

[Invoice Report (5).xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%204/Invoice%20Report%20(5).xlsx)

Una vez creada una tabla debemos tener en cuenta que los datos contenidos en ellas, al usar funciones, se ubican con referencias estructuradas. Por ejemplo:

| Con tabla | =SUMAR.SI(tbl_MC[Location],$A8,tbl_MC[Amount]) |
| --- | --- |

| Con rango y tabla | =SUMAR.SI(Location,$A8,tbl_MC[Over Due By]) |
| --- | --- |

| Con rango | =SUMAR.SI(Location,A8,Amount_Paid) |
| --- | --- |

| Con ref. celda | =SUMAR.SI(I5:I88,'Recon Analysis'!A8,J5:J88) |
| --- | --- |


Supongamos que quiero agregar un pago nuevo a la hoja “MC Invoice Report”, pero correspondiente a una nueva region llamada Perth. Debería agregar “Perth” a la hoja “Recon Analysis” donde tengo el resumen por zona. Armo una tabla (tbl_Region) para que los datos se actualicen solos. 

[Invoice Report (6).xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%204/Invoice%20Report%20(6).xlsx)

Si deseo actualizar los registros antiguos con nuevos datos, basta con eliminar las filas correspondientes y agregar las nuevas. De esta manera, todas las operaciones de limpieza de datos previamente hechas se aplicarán automáticamente.

![imagen1](https://user-images.githubusercontent.com/110131341/226633898-f846a296-e738-46c0-8204-a0ad757b700a.png)



#### Condicionales

- IF encadenados
- Funciones LOOKUP (categorización y match), INDEX, MATCH.

En la hoja “Recon Analysis” puede observarse una diferencia entre el pago según el proveedor y según nuestro sistema.

En la planilla del proveedor hay unos registros que corresponden a créditos (Type = CR). Estos valores deben restarse, puede que la diferencia se deba a eso. Para solucionarlo usamos un condicional tal que cada vez que aparezca un crédito, esa factura se reste del total.

| Celda | Fórmula |
| --- | --- |
| tbl_Supplier[[#Encabezados],[$ Amount]] | =VALOR(SUSTITUIR(SUSTITUIR(F2,"S",""),EXTRAE(F2,2,1),""))*SI([@Type]="CR",-1,1) |


<br><br>
Supongamos que ahora me dicen que los recargos por vencimiento son condicionales según la región y según la cantidad de días de retraso. Si se trata de Melbourne, se cobra $5 por cada día de vencimiento. Si se trata de Sydney, usamos la nueva tabla de penalidades:

| Over Due by | Charge|
| --- | --- |
| 0 | $0.00 |
| 1 | $2.25 |
| 5 | $5.50 |
| 10 | $10.80 |
| 15 | $25.90 |

En resumen, quedaria el siguiente esquema 

[![](https://mermaid.ink/img/pako:eNptj8EKgkAQhl9lGDoYKHTp4iGpzGMERRfXw-COJelubGsh5ru3ZYSH5jQM3_8xf4e5lowhFpV-5GcyFg6xUOBmme7opOHOKi-lhiiDIFg8E6pu-gkrbzb9h32pIxtJko0j196-lYpbiIbAkFqPZXGaVGTBkGXQBUzm2RgauTbplh-wY0WVbT989nOijzWbmkrp2nTvi0B75poFhm6VZC4CheodR43V-1blGFrTsI_NVTpVXNLJUI1h4b7i_gXy71og?type=png)](https://mermaid.live/edit#pako:eNptj8EKgkAQhl9lGDoYKHTp4iGpzGMERRfXw-COJelubGsh5ru3ZYSH5jQM3_8xf4e5lowhFpV-5GcyFg6xUOBmme7opOHOKi-lhiiDIFg8E6pu-gkrbzb9h32pIxtJko0j196-lYpbiIbAkFqPZXGaVGTBkGXQBUzm2RgauTbplh-wY0WVbT989nOijzWbmkrp2nTvi0B75poFhm6VZC4CheodR43V-1blGFrTsI_NVTpVXNLJUI1h4b7i_gXy71og)

[Invoice Report (7).xlsx](https://github.com/palomachcc/Excel/raw/main/Parte%205/Invoice%20Report%20(7).xlsx)
