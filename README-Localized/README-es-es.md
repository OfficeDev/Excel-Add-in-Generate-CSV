# Ejemplo del complemento del panel de tareas del generador de CSV para Excel 2016

_Se aplica a: Excel 2016_

Este complemento del panel de tareas muestra cómo crear una tabla a partir de una lista de nombres de columna con las API de JavaScript de Excel 2016. Hay dos tipos: editor de código y Visual Studio.

![Ejemplo de generador de CSV](../Images/ScreenCap1.PNG)

## Pruébelo
### Versión del editor de código

La forma más sencilla de implementar y probar el complemento consiste en copiar los archivos en un recurso compartido de red.

1.  Cree una carpeta en un recurso compartido de red (por ejemplo, \\\MiRecursoCompartido\Excel_CSV_Generator) y, después, copie todos los archivos en la carpeta del Editor de código. 
2.  Edite el elemento <SourceLocation> del archivo de manifiesto para que apunte a la ubicación del recurso compartido creada en el paso 1. 
3.  Copie el manifiesto (TeacherCSVGenerator.xml) en un recurso compartido de red (por ejemplo, \\\MiRecursoCompartido\MisManifiestos).
4.  Agregue la ubicación del recurso compartido que contiene el manifiesto como un catálogo de aplicaciones de confianza en Excel.

    a. Inicie Excel y abra una hoja de cálculo en blanco.  
    
    b. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
    
    c. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
    
    d. Elija **Catálogos de complementos de confianza**.
    
    e. En el cuadro **URL de catálogo**, escriba la ruta de acceso al recurso compartido de red que creó en el paso 3 y luego elija **Agregar catálogo**.
    
   f. Active la casilla **Mostrar en el menú** y elija **Aceptar**. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Office. 
        
5.  Pruebe y ejecute el complemento. 

    a. En la pestaña **Insertar** de Excel 2016, elija **Mis complementos**. 
    
    b. En el cuadro de diálogo **Complementos de Office**, seleccione **Carpeta compartida**.
    
    c. Elija **Ejemplo de lista de clase para profesores en CSV** > **Insertar**. El complemento se abre en el panel de tareas y crea el CSV con la lista de clase en la hoja activa, tal como se muestra en la captura de pantalla. 
      
   ![Ejemplo de rastreador de presupuestos universitarios](../Images/ScreenCap2.PNG) 

    d. Elija un servicio de gestión de aulas.
    
    e. Haga clic en el botón Crear lista para insertar una lista vacía en la hoja de cálculo activa.  
    
      ![Ejemplo de seguimiento del presupuesto universitario](../Images/ScreenCap3.PNG) 
      
    f. Haga clic en el botón Ayuda con la exportación de Excel para obtener más información sobre cómo exportar la hoja de cálculo como un archivo .csv.  
  
    
### Versión de Visual Studio
1.  Copie el proyecto en una carpeta local y abra TeacherCSVGenerator.sln en Visual Studio.
2.  Pulse F5 para crear e implementar el complemento de ejemplo. Excel se inicia y se abre el complemento en un panel de tareas a la derecha de una hoja de cálculo en blanco, como se muestra en la siguiente captura de pantalla. 
        
  ![Ejemplo de generador de CSV de Excel](../Images/ScreenCap1.PNG) 

3.  Seleccione un servicio de administración de clases en línea de la lista desplegable.
4.  Agregue una tabla de lista de alumnos con el botón **Make roster** (Hacer lista) y vea la tabla creada en la hoja de cálculo activa.

  ![Ejemplo de rastreador de presupuestos universitarios](../Images/ScreenCap3.PNG) 
5.  Rellene las celdas en las filas bajo el encabezado de la tabla para agregar alumnos a la lista.
6.  Use la característica de exportar en Excel para guardar la hoja de cálculo como un archivo .csv. Este archivo está en el formato correcto para importar en el servicio de su elección.


### Obtener más información

Las API de JavaScript de Excel tienen mucho que ofrecer para el desarrollo de complementos. A continuación se muestran algunos de los recursos disponibles. 

1.  [Introducción a la programación de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de fragmentos de código para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Ejemplos de código de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Referencia de la API de JavaScript de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Compilar el primer complemento de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)

