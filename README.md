# Excel ![Badge en Desarollo](https://img.shields.io/badge/STATUS-EN%20DESAROLLO-green)

## :hammer:Repaso de algunas herramientas básicas de excel para limpieza y orden de datos.

- `Parte 1`: Cleaning and manipulating text. Concatenar o unir textos, extraccion de partes. Removing and replacing text characters
- `Parte 2`: Format data. Varias funciones utiles para cuando laburas con fechas, calculos con dias habiles o vacaciones etc.
- `Parte 3`: Cell referencing and naming; Creating named ranges; Managing named ranges; Calculations with named ranges; and Automating processes with named ranges.
- `Parte 4`: Tablas, referencias estructuradas.
- `Parte 5`: Operaciones logicas con IF, AND, OR. VLOOKUP, XLOOKUP, INDEX, MATCH.
- `Parte 6`: Tablas dinámicas. 
-------
A la hora de trabajar con datos es necesario que estén limpios y ordenados para asegurarnos de que sean coherentes. Esto también ayuda a reducir la posibilidad de errores y mejora la eficiencia a la hora de analizarlos. De esta manera es mas fácil detectar patrones o tendencias y poder así tomar decisiones informadas.


### Ejemplo 1. Reporte de pagos

Contexto: Recibo un archivo (Hoja Supplier Invoice Statement) de un proveedor con los pagos realizados en el mes (Abril). La idea es ingresar esos datos a mi sistema, verificando que todo coincide y esta en orden (MC Invoice Report). Todo se descarga y se realiza en una hoja de excel. 
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
(Ctrl+T)

Una vez creada la tabla podemos ordenar y filtrar los datos según lo que necesitemos.

![image2](https://user-images.githubusercontent.com/110131341/226460512-8cab21e6-d66a-4dcd-8b11-7da63f737061.png)

Tambien podemos agregar slicers, con la opción de Segmentación de datos, que permiten insertar filtros de una manera interactiva. Además del filtrado rápido, las segmentaciones de datos también indican el estado de filtrado actual, lo que facilita la comprensión de lo que se muestra exactamente en ese momento.

![image3](https://user-images.githubusercontent.com/110131341/226460687-01678920-5773-4af3-b91c-35d0587440dd.png)

[Invoice Report (5).xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%204/Invoice%20Report%20(5).xlsx)

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

[Invoice Report_ejemplo_Perth.xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%205/Invoice%20Report_ejemplo_Perth.xlsx)

Si deseo actualizar los registros antiguos con nuevos datos, basta con eliminar las filas correspondientes (no elimina la tabla) y agregar las nuevas. De esta manera, todas las operaciones de limpieza de datos previamente hechas se aplicarán automáticamente.

![imagen1](https://user-images.githubusercontent.com/110131341/226633898-f846a296-e738-46c0-8204-a0ad757b700a.png)

En caso de querer eliminar la tabla, los datos se mantienen. Para eso usamos la opcion "Convertir en rango".

#### Operaciones lógicas y condicionales

- IF encadenados
- Funciones LOOKUP (categorización y match), INDEX, MATCH.

En la hoja “Recon Analysis” puede observarse una diferencia entre el pago total según el proveedor y según nuestro sistema.

En la planilla del proveedor hay unos registros que corresponden a créditos (Type = CR). Los totales correspondientes a estos creditos deben ser negativos, puede que la diferencia se deba a esto. Para solucionarlo usamos un condicional tal que cada vez que aparezca un crédito, esa factura se reste del total.

| Celda | Fórmula |
| --- | --- |
| tbl_Supplier[[#Encabezados],[$ Amount]] | =VALOR(SUSTITUIR(SUSTITUIR(F2,"S",""),EXTRAE(F2,2,1),""))*SI([@Type]="CR",-1,1) |

Es decir, si en la columna Type aparece CR, multiplica el amount por -1.
La diferencia se resuelve.

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

[Invoice Report (7).xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%205/Invoice%20Report%20(7).xlsx)

El condicional IF (SI en español) solo puede manejar dos resultados. Una opción es usar varios “IF” encadenados. Si quiero agregar mas de una prueba lógica puedo usar OR (devuelve verdadero si se cumple alguna de las condiciones) o AND (devuelve verdadero si se cumplen todas las condiciones).

Otra opción es combinar el condicional IF (SI) con VLOOKUP (BUSCARV).

Para este caso, primero ubicamos la penalidad correspondiente, teniendo en cuenta el diagrama previo:

| Celda | Formula |
| --- | --- |
| O5 “New Penalty” | =SI([@Location]="Sydney";BUSCARV([@[Over Due By]];$Q$22:$R$26;2);5) |

Estoy indicando lo siguiente: si la ubicación coincide con Sydney, devolver el resultado de la funcion BUSCARV, de lo contrario, el resultado es 5. 

Lo que hace la funcion BUSCARV en este caso es evaluar los días de retraso y devolver la penalidad correspondiente según los días vencidos. Por ejemplo, entre 10 y 15 días se cobra $10.80.

Notas:

-BUSCARV (dias de retraso, rango de datos donde busco la info, columna que deseo devolver del rango de datos).

-El rango de datos debe estar ordenado de menor a mayor.

-Compara el primer parámetro con la primer columna del rango de datos indicado. En el tercer parametro especificamos la columna que contiene los posibles resultados.

-La “V” se refiere a vertical. Esta funcion sirve solamente si la info esta organizada verticalmente. Para info horizontal usamos HLOOKUP.

Otro uso que le podríamos dar a la funcion es el de verificar que el valor de cada factura, según el proveedor, coincide con los valores que tengo en mi sistema. 

El “ID” de cada factura es el numero de pago. Armo una columna nueva donde evaluo, según el numero de pago, cuales son los valores de cada factura y si hay diferencia entre ellos.

| Verificación valores (celda U2 en “Supplier Invoice Statement”) | = Q2-BUSCARV(tbl_Supplier[@[Payment No]];tbl_MC;10;FALSO) |
| --- | --- |

Tambien existe la funcion XLOOKUP (BUSCARX en español) que tiene mas variedad de opciones para la búsqueda pero no se encuentra en todas las versiones de Excel.

[Invoice Report (8).xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%205/Invoice%20Report%20(8).xlsx)

Comparando los valores te das cuenta que hay una diferencia, los valores estan bien pero uno de los ID se repite y eso lleva a un error. Corrigiendo eso, queda todo en orden.
[Invoice Report Final.xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%205/Invoice%20Report%20Final.xlsx)


### Ejemplo 2. Buscador 

Tengo un archivo con la poblacion respectiva de diferentes paises. 

[Country_Population_Inicial.xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%205/Country_Population_Inicial.xlsx)

La idea es poder ingresar el pais en una celda y que me devuelva automaticamente la poblacion correspondiente. Para esto son utiles las funciones INDICE (INDEX) y COINCIDIR (MATCH).

-La función COINCIDIR busca un elemento determinado en un intervalo de celdas y después devuelve la posición relativa de dicho elemento en el rango. Por ejemplo, si el rango A1:A3 contiene los valores 5, 25 y 38, la fórmula =COINCIDIR(25,A1:A3,0) devuelve el número 2, porque 25 es el segundo elemento del rango.
-La función INDICE nos permite encontrar un valor en un rango (matriz) especificando el valor de la posición del dato buscado a través de la fila y la columna. Por ejemplo, si quiero buscar la poblacion de Argelia que se encuentra tercera en la columna "Population", escribo =INDICE(Population,3) =INDICE(columna,fila).


Lo primero que hago es nombrar rangos de datos, uno para cada columna. (ctrl+shift+f3)
Para la celda B4 uso lo que vimos de validacion de datos, de modo que me quede una lista desplegable de los paises (rango "Ctry").

[Country_Population_Final.xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%205/Country_Population_Final.xlsx)

En la primer hoja del archivo busca solamente por pais. En la segunda hoja agrego mas opciones, es como quedaria el trabajo final.

### Ejemplo 3. Ventas
[Sales.xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%206/Sales.xlsx)

En este archivo tengo las ventas del 2020 de ciertos productos de una empresa de alimentos. La idea es hacer un analisis rápido usando tablas dinamicas. 

Lo primero que puedo hacer es armar una tabla con los datos y agregar una nueva columna llamada "Month", con el mes respectivo pero en texto. Con "ctrl + A" selecciono todos los datos de manera rapida.

Armo una tabla dinamica para analizar las ventas por mes. Conviene tenerla en una hoja aparte para no mezclar los datos. En este caso la llamamos "Sales by month". Podemos agregar un grafico de linea. Con esto vemos que no hay una tendencia particular a lo largo del año pero si existe un pico de ventas en el mes de junio.

Tambien podemos armar otra tabla dinamica para analizar las ventas por persona. En este caso uso un grafico de barras. Siempre se puede acomodar el formato, ordenar los datos de mayor a menor, colores, tipo de moneda, etc.

[Sales (2).xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%206/Sales%20(2).xlsx)

Los valores pueden representarse de diferentes maneras. Por ejemplo, si agrego una tabla dinamica para los items vendidos segun la categoria, podria representar los valores en porcentaje respecto al total (un poco mas representativo que las cantidades por unidad). Selecciono los datos a representar-->clic derecho--> mostrar valores como.

![image](https://user-images.githubusercontent.com/110131341/235736948-db9d1a33-0016-4f95-b83d-ad7a425f0d41.png)

La info que voy resumiendo puede ser plasmada en un dashboard. Para esto armamos una nueva hoja. Unificamos el fondo, seleccionando todas las celdas haciendo clic en el extremo superior izquierdo y asignandole un color de fondo.

![image](https://user-images.githubusercontent.com/110131341/235735103-f65da6fd-8792-4963-979a-6c0ca1e59980.png)

Por ultimo copiamos y pegamos los graficos o tablas que sean necesarios. 

![image](https://github.com/palomachcc/Excel/assets/110131341/e68c579b-1b22-45e2-8b1d-5c5f0373dfed)

[Sales (3).xlsx](https://github.com/palomachcc/Excel/blob/main/Parte%206/Sales%20(3).xlsx)


