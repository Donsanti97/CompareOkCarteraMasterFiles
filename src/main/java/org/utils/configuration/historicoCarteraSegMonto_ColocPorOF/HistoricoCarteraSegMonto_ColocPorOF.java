package org.utils.configuration.historicoCarteraSegMonto_ColocPorOF;

import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.utils.FunctionsApachePoi.*;
import static org.utils.MethotsAzureMasterFiles.runtime;

public class HistoricoCarteraSegMonto_ColocPorOF {
    //110 hojas

    public static void carteraBruta(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void hoja1(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja2(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja3(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja4(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja5(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja6(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja7(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja8(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja9(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja10(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja11(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja12(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja13(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja14(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja15(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja16(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja17(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja18(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja19(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja20(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja21(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja22(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja23(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja24(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja25(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja26(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja27(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja28(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja29(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja31(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja32(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja33(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja34(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja35(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja36(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja37(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja38(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja39(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja40(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja41(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja42(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja43(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja44(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja45(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja46(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja47(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja48(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja49(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja50(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja51(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja52(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja53(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja54(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja55(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja56(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja57(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja58(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja59(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja60(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja61(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja62(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja63(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja64(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja65(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja66(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja67(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja68(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja69(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja70(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja71(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja72(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja73(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja74(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja75(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja76(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja77(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja78(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja79(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja80(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja81(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja82(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja83(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja84(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja85(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja86(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja87(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja88(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja89(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja90(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja91(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja92(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja93(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja94(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja95(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja96(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja97(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja98(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja99(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void hoja100(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja101(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja102(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja103(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja104(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja105(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja107(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja108(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja109(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }
    public static void hoja110(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "CONCUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONCUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("----------------------");
        }
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

}
