package org.utils.configuration.historicoCarteraConsumoPorOF;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static org.utils.MethotsAzureMasterFiles.*;
import static org.utils.FunctionsApachePoi.*;

public class HistoricoCarteraConsumoPorOF {
    //34 Hojas
    public static void deleteTempFile(String tempFile) {
        eliminarExcel(tempFile, 5);
    }

    public static void configuracion(String masterFile) {

        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        /*JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        String masterFile = getDocument();*/
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola el número del mes y año de corte del archivo OkCartera sin espacios (Ejemplo: 02/2023 (febrero/2023))");
        String mesAnoCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "A continuación se creará un archivo temporal " +
                "\n Se recomienda seleccionar la carpeta \"Documentos\" para esta función...");
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";




        try {

            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitMinutes(5);

            carteraBruta(okCartera, masterFile, azureFile, fechaCorte, "Cartera Bruta", tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "0 Dias", 0, 0, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "1 - 7 Dias", 1, 7, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "7 - 15 Dias", 8, 15, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "16 - 30 Dias", 16, 30, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "31 - 60 Dias", 31, 60, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "61 - 90 Dias", 61, 90, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "91 - 120 Dias", 91, 120, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "121 - 150 Dias", 121, 150, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "151 - 180 Dias", 151, 180, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "181 - 360 Dias", 181, 360, tempFile);
            waitSeconds(5);


            diasDeMoraDias(okCartera, masterFile, azureFile, fechaCorte, "> 361 Dias", 361, 5000, tempFile);
            waitSeconds(5);


            calificacion(okCartera, masterFile, azureFile, fechaCorte, "A", "A", tempFile);
            waitSeconds(5);


            calificacion(okCartera, masterFile, azureFile, fechaCorte, "B", "B", tempFile);
            waitSeconds(5);


            calificacion(okCartera, masterFile, azureFile, fechaCorte, "C", "C", tempFile);
            waitSeconds(5);


            calificacion(okCartera, masterFile, azureFile, fechaCorte, "D", "D", tempFile);
            waitSeconds(5);


            calificacion(okCartera, masterFile, azureFile, fechaCorte, "E", "E", tempFile);
            waitSeconds(5);


            reEstCapital(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital", tempFile);
            waitSeconds(5);


            reEstCapital(okCartera, 0, 150, masterFile, azureFile, fechaCorte, "Re_Est Capital < = 150", tempFile);
            waitSeconds(5);


            reEstCapital(okCartera, 151, 5000, masterFile, azureFile, fechaCorte, "Re_Est Capital > 150", tempFile);
            waitSeconds(5);


            reEstNCreditos(okCartera, masterFile, azureFile, fechaCorte, "Re_Est N° Creditos", tempFile);
            waitSeconds(5);


            nCreditosVigentes(okCartera, masterFile, azureFile, fechaCorte, "N° Creditos Vigentes", tempFile);
            waitSeconds(5);

            clientesComercial(okCartera, masterFile, azureFile, fechaCorte, "Clientes_Comercial", tempFile);
            waitSeconds(5);

            colocacionComercial(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Colocación_Comercial", tempFile);
            waitSeconds(5);

            nCreditoComercial(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "N° De Créd Comercial", tempFile);
            waitSeconds(5);

            colocacionPromComercial(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Colocación Prom Comercial", tempFile);
            waitSeconds(5);

            comercialPercentil05(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Comercial Percentil 0.5", tempFile);
            waitSeconds(5);

            comercialPercentil08(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Comercial Percentil 0.8", tempFile);
            waitSeconds(5);

            comercialPzoPerc05(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Comercial_Pzo_Perc_0.5", tempFile);
            waitSeconds(5);

            comercialPzoProm(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Comercial_Pzo_Prom", tempFile);
            waitSeconds(5);

            JOptionPane.showMessageDialog(null, "Espere un momento la última hoja está siendo analizada. \n Por favor de clic en Ok para continuar...");
            waitSeconds(5);

            comercialPzoPerc08(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Comercial_Pzo_Perc_0.8", tempFile);
            waitSeconds(5);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }


    }

    public static void carteraBruta(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, valorInicio, valorFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void diasDeMoraDias(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangoDesde, int rangoHasta, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {

            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoDiasDeMora = "dias_de_mora";
            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            //int rangoDesde = 361;
            //int rangoHasta = 5000L;

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, campoFiltrar, valorInicio, valorFin, campoDiasDeMora, rangoDesde, rangoHasta);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void calificacion(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoCalificacion = "calificacion";
            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSS(sheet, headers, campoFiltrar, valorInicio, valorFin, campoCalificacion, calificacion, calificacion);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void reEstCapital(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String reEstCapital = "re_est";
            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, campoFiltrar, valorInicio, valorFin, reEstCapital, 1, 1);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void reEstCapital(String okCarteraFile, int diasMoradesde, int diasMoraHasta, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String reEstCapital = "re_est";
            String diasDeMora = "dias_de_mora";

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, reEstCapital, 1, 1, diasDeMora, diasMoradesde, diasMoraHasta);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void reEstNCreditos(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, valorInicio, valorFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void nCreditosVigentes(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {

            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("linea", "Cliente");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, valorInicio, valorFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void clientesComercial(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {

            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, valorInicio, valorFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void colocacionComercial(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void nCreditoComercial(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void colocacionPromComercial(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void comercialPercentil05(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), 50);
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }


    public static void comercialPercentil08(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), 80);
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    /*---------------------------------------------------------------------------------------------------------------------*/
    public static void comercialPzoProm(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    //Mertodos a los que hay que hacerle un método aparte en la tabla dinámica para hallar el porcentaje 50%
    public static void comercialPzoPerc05(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);


            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), 50);
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    //Mertodos a los que hay que hacerle un método aparte en la tabla dinámica para hallar el porcentaje 80%
    public static void comercialPzoPerc08(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "modalidad";
            String valorInicio = "CONSUMO"; // Reemplaza con el valor de inicio del rango
            String valorFin = "CONSUMO"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSD(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), 80);
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

}
