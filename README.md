<b>Objetivo:</b>
<br>Comparar los valores dados en un archiv Ok Cartera, frente a los valores encontrados en un archivo Maestro en una fecha de corte específica

<b>Precondiciones:</b>
<br>Tener como mínimo JAVA versión 11 instalado y configurado en sus variables de entorno
Tener un computador como una capacidad mínima en RAM de 32GB

<b>Paso a paso:</b>
<br>Para ejecutar el paso a paso debe tener claros los archivos que va a comparar, ya que el programa no diferenciará si usted 
elije dos archivos distintos el uno del otro. El programa está diseñado únicamente para compara los valores que le indique a través del proceso.
Para comenzar el proceso se le compartirá un archivo comprimido con los siguientes elementos:
1. Dos archivos ejecutables, uno con extensión .SH(Para sistema operativo Linux), y .BAT(Para sistema operativo Windows)
2. Una carpeta donde se aloja un ejecutable .JAR
3. Una carpeta "documentos" que contendrá dos carpetas llamadas "ArchivosAzure" y "ArchivosMaestro", y "OkCarteraFiles" carpetas que de las cuales se le recomienda hacer
   uso para alojar los archivos a analizar.
<br><b>Recomendación:</b>Para mas comodidad descomprima la carpeta en el area de Documentos de su computadora.

<b>Durante la ejecución:</b>
1. Deberá iniciar una consola CMD ejecutándola como administrador. Esto es posible buscando en su barra de tareas "CMD", y dándo clic derecho sobre la primera opción que aparezca
   en la búsqueda, y seleccionando la opción "Ejecutar como administrador".
   
   ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/e2e5dceb-170c-45c1-a0aa-3abc3d44a003)

3. a continuación deberá navegar a través de las carpetas para llegar a la ubicación del archivo ejecutable (.SH(linux) .BAT(Windows)).
   <br><b>Comandos:</b>
   <ul>
     <li><b>"cd ..":</b>Ingresa en una carpeta anterior a la actual</li>
     <li><b>"cd 'nombre_carpeta/directorio'":</b> Ingresa a una carpeta o ruta que le sea indicado según necesidad</li>
   </ul>
<h2>Regresar a la carpeta hasta estar en la raíz (carpeta C:)</h2>

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/fef2e1de-5b09-42b2-ab67-295a6497eff1)

<h2>Ingresar a la carpeta donde se ubica el ejecutable</h2>

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/cfe15f69-b045-44f1-b74f-495a8db5c9c8)
![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/4a587570-f66a-45a7-8f2e-2868d4c8aee3)

4. A continuación deberá escribir el nombre del archivo ejecutable y dar enter. Esto dará paso para el inicio de la ejecución del aplicativo

  ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/2032ed1d-03ae-48e9-a350-3bf44ca4c083)


5. Seleccione el archivo Maestro según indicación. Para esto deberá buscar entre los directorios en la ventana emergente hasta encontrar el que desea analizar
   
   ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/4b716c23-acac-4557-b2ef-83eb96239878)

   ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/25371bcd-9e4d-42d7-aa53-0da602aa9fd1)


6. Seleccione el archivo Azure según indicación. Al igual que el paso anterior deberá buscar entre los directorios hasta encontrar el archivo que se iguale al archivo a analizar

   ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/82a7c5b0-5c6a-4f39-afc2-854e9cd7ec4b)

   
6. A continuación deberá seleccionar el archivo Ok Cartera que será usado para filtrar la información que será comparada contra el archivo Maestro

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/92caf5d6-0441-46aa-bc82-53360012bf7c)

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/47a2cf99-bcfa-48f6-b58b-b406cebf2628)


   
7. El paso siguiente, es seleccionar el mes y año de corte de la cual desdea extraer la información

   ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/40732d62-0f72-4b90-998f-b459a2ba4fba)

  ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/75fee946-1c5b-42d1-b597-de3920117522)


10. A continuación se le pedirá que seleccione una carpeta aleatorio donde será alojado un archivo temporal para uso exclusivo de la ejecución. Se recomienda no mover ni eliminar tal archivo.

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/57147511-78d2-4846-903b-376d61b03cb2)

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/3f7ec2f3-725a-4766-bccf-18edd0970bcd)


12. Los siguientes pasos serán durante la ejecución. Por cada análisis de hojas el programa le pedirá que ingrese el encabezado "Código", y el encabezado que contenga los valores que analizará, y que será usado para comparar contra los que se encuentren en el archivo Ok Cartera seleccionado.

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/874c4dcf-c7dc-4986-b68b-f1a53f2e10d7)

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/4873272b-6bb1-462b-8ec8-af6a51e3f3fe)

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/733809c2-86d8-483e-a99a-8b873a8d095b)

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/99259062-8ca8-429d-aac5-69f792fd6185)
    
<br><b>Nota: Si no hay encabezado de valores que se ajuste a lo que necesitas analizar, deberás seleccionar la opción "Ninguno", el programa preguntará si deseas analizar otro encabezado diferente. En caso de que la respuesta sea "NO", se mostrará por pantalla que la información está incompleta y que esa hoja en específico no podrá ser analizada, y continuará con la siguiente. Por el contrario si se selecciona "SI" le pedirá que seleccione el encabezado que desea comparar en valores</b>

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/54375483-fa19-4614-a000-2ac5c4115cfa)

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/8773bfff-e7a8-4f26-b126-a000364f4cc1)

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/f653e865-6817-4a87-8038-cdd1cf282f9f)

![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/fff5e900-968e-4760-a2f9-8bf5cbbe4a5b)


<br><b>Nota: El paso anterior se ejecutará el número de veces, equivalente al número de hojas a analizar en el archivo Maestro</b>

13. Cuando finalice el proceso, en la consola se mostrarán las carpetas donde fueron alojados las igualdades y diferencias del análisis. Estas carpetas serán creadas automáticamente como "errores" y "messages" en la carpeta de los archivos Maestro. Estos se nombrarán con el nombre del archivo maestro analizado, la fecha y hora de la finalización del análisis.

    ![image](https://github.com/Donsanti97/CompareOkCarteraMasterFiles/assets/47354432/ace43c82-e4f5-4bd8-8036-bdaee158a253)
