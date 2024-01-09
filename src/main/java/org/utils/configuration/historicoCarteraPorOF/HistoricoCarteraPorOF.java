package org.utils.configuration.historicoCarteraPorOF;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static org.utils.FunctionsApachePoi.*;
import static org.utils.MethotsAzureMasterFiles.*;
import static org.utils.configuration.GetMasterAnalisis.*;


public class HistoricoCarteraPorOF {
    //64 Hojas

    public static boolean isEqual(String azureFile){
        boolean isEqual = false;
        File aFile = new File(azureFile);
        if (aFile.getName().toLowerCase().contains("cartera por oficina")){
            isEqual = true;
        }
        return isEqual;
    }

    public static void configuracion(String masterFile) {

        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        while (!isEqual(azureFile)){
            errorMessage("El archivo AZURE no es el indicado para el análisis." +
                    "\n \n Por favor seleccione el archivo correspondiente a: " + new File(masterFile).getName());
            azureFile = getDocument();
        }
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola el número del mes y año de corte del archivo OkCartera sin espacios (Ejemplo: 02/2023 (febrero/2023))");
        String mesAnoCorte = showMonthYearChooser();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = showDateChooser();

        while (azureFile == null || okCartera == null || mesAnoCorte == null || fechaCorte == null){
            errorMessage("Alguno de los items requeridos anteriormente no fue seleccionado." +
                    "\n Por favor seleccione nuevamente los items requeridos.");
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
            azureFile = getDocument();
            while (!isEqual(azureFile)){
                errorMessage("El archivo AZURE no es el indicado para el análisis." +
                        "\n \n Por favor seleccione el archivo correspondiente a: " + new File(masterFile).getName());
                azureFile = getDocument();
            }
            JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
            okCartera = getDocument();
            JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola el número del mes y año de corte del archivo OkCartera sin espacios (Ejemplo: 02/2023 (febrero/2023))");
            mesAnoCorte = showMonthYearChooser();
            JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
            fechaCorte = showDateChooser();
        }

        JOptionPane.showMessageDialog(null, "A continuación se creará un archivo temporal " +
                "\n Se recomienda seleccionar la carpeta \"Documentos\" para esta función...");
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";

        try {
            waitSeconds(10);
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitSeconds(5);

            List<String> machSheets = machSheets(azureFile, masterFile);

            //waitMinutes(8);

            carteraBruta(okCartera, masterFile, azureFile, fechaCorte, "Cartera Bruta", tempFile, machSheets);

            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "0 Dias", 0, 0, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "1 - 7 Dias", 1, 7, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "8 - 15 Dias", 8, 15, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "16 - 30 Dias", 16, 30, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "31 - 60 Dias", 31, 60, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "61 - 90 Dias", 61, 90, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "91 - 120 Dias", 91, 120, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "121 - 150 Dias", 121, 150, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "151 - 180 Dias", 151, 180, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "181 - 360 Dias", 181, 360, tempFile, machSheets);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "> 361 Dias", 361, 5000, tempFile, machSheets);

            calificacion(okCartera, masterFile, azureFile, fechaCorte, "A", "A", tempFile, machSheets);
            calificacion(okCartera, masterFile, azureFile, fechaCorte, "B", "B", tempFile, machSheets);
            calificacion(okCartera, masterFile, azureFile, fechaCorte, "C", "C", tempFile, machSheets);
            calificacion(okCartera, masterFile, azureFile, fechaCorte, "D", "D", tempFile, machSheets);

            reEstCapital(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital", tempFile, machSheets);

            reEstCapital(okCartera, 0, 30, masterFile, azureFile, fechaCorte, "Re_Est Capital < = 30", tempFile, machSheets);
            reEstCapital(okCartera, 31, 5000, masterFile, azureFile, fechaCorte, "Re_Est Capítal > 31", tempFile, machSheets);

            reEstNCreditos(okCartera, masterFile, azureFile, fechaCorte, "Re_Est N° Creditos", tempFile, machSheets);
            nCreditosVigentes(okCartera, masterFile, azureFile, fechaCorte, "N° Creditos Vigentes", tempFile, machSheets);

            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital-A", "A", tempFile, machSheets);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital-B", "B", tempFile, machSheets);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital-C", "C", tempFile, machSheets);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital-D", "D", tempFile, machSheets);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital-E", "E", tempFile, machSheets);

            reestructuradosCapitalOfcPlazosProm(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_Plazos_Prom", tempFile, machSheets);

            reestructuradosCapitalOfcPlazosMin(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_Plazos_Min", tempFile, machSheets);

            reestructuradosCapitalOfcPlazosMax(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_Plazos_Max", tempFile, machSheets);

            mora1raCuotaMontoOfc(okCartera, masterFile, azureFile, fechaCorte, "Mora-1raCuota_Monto_Ofc", tempFile, machSheets);

            mora1raCuotaCantOfc(okCartera, masterFile, azureFile, fechaCorte, "Mora-1raCuota_Ofc_Cant", tempFile, machSheets);

            provisiones(okCartera, masterFile, azureFile, fechaCorte, "Provisiones", tempFile, machSheets);

            clientes(okCartera, masterFile, azureFile, fechaCorte, "Clientes", tempFile, machSheets);

            colocacion(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Colocación", tempFile, machSheets);

            credPromColocacion(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Cred Prom Colocación", tempFile, machSheets);

            colocacionPercentil(okCartera, masterFile, azureFile, 50, mesAnoCorte, fechaCorte, "Colocación Percentil 0.5", tempFile, machSheets);

            colocacionPercentil(okCartera, masterFile, azureFile, 80, mesAnoCorte, fechaCorte, "Colocación Percentil 0.8", tempFile, machSheets);

            cartera(okCartera, masterFile, azureFile, fechaCorte, "Cartera >=1", 1, 5000, tempFile, machSheets);
            cartera(okCartera, masterFile, azureFile, fechaCorte, "Cartera >30", 31, 5000, tempFile, machSheets);

            carteraBE(okCartera, masterFile, azureFile, fechaCorte, "Cartera B-E", "B", tempFile, machSheets);

            plazoPercentil(okCartera, masterFile, azureFile, 50, mesAnoCorte, fechaCorte, "Plazo Percentil 0.5", tempFile, machSheets);

            plazoPercentil(okCartera, masterFile, azureFile, 80, mesAnoCorte, fechaCorte, "Plazo Percentil 0.8", tempFile, machSheets);

            pzoPromOfc(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Pzo_Prom_Oficinas", tempFile, machSheets);

            carteraMoraCronograma(okCartera, masterFile, azureFile, fechaCorte, "Cartera Mora Cronograma >8", 9, 5000, tempFile, machSheets);
            carteraMoraCronograma(okCartera, masterFile, azureFile, fechaCorte, "Cartera Mora Cronograma =0", 0, 0, tempFile, machSheets);

            clientesMoraCronograma(okCartera, masterFile, azureFile, fechaCorte, "Clientes Mora Cronograma =0", 9, 5000, tempFile, machSheets);

            modificadosCapital(okCartera, masterFile, azureFile, fechaCorte, "Modificados Capital", tempFile, machSheets);

            modificados(okCartera, masterFile, azureFile, fechaCorte, "Modificados < = 30", 0, 30, tempFile, machSheets);
            modificados(okCartera, masterFile, azureFile, fechaCorte, "Modificados > 30", 31, 5000, tempFile, machSheets);

            modificadosNCreditos(okCartera, masterFile, azureFile, fechaCorte, "Modificados N° Créditos", tempFile, machSheets);

            modificadosCapital(okCartera, masterFile, azureFile, fechaCorte, "Modificados Capital-A", "A", tempFile, machSheets);
            modificadosCapital(okCartera, masterFile, azureFile, fechaCorte, "Modificados Capital-B", "B", tempFile, machSheets);
            modificadosCapital(okCartera, masterFile, azureFile, fechaCorte, "Modificados Capital-C", "C", tempFile, machSheets);
            modificadosCapital(okCartera, masterFile, azureFile, fechaCorte, "Modificados Capital-D", "D", tempFile, machSheets);
            modificadosCapital(okCartera, masterFile, azureFile, fechaCorte, "Modificados Capital-E", "E", tempFile, machSheets);

            modificadoCapitalOfcPlazosProm(okCartera, masterFile, azureFile, fechaCorte, "Modificado Capital_Plazos_Prom", tempFile, machSheets);

            modificadoCapitalOfcPlazosMin(okCartera, masterFile, azureFile, fechaCorte, "Modificado Capital_Plazos_Min", tempFile, machSheets);

            modificadoCapitalOfcPlazosMax(okCartera, masterFile, azureFile, fechaCorte, "Modificado Capital_Plazos_Max", tempFile, machSheets);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            logWinsToFile(masterFile, coincidencias);
            logErrorsToFile(masterFile, errores);

            deleteTempFile(tempFile);
        } catch (HeadlessException | ParseException | IOException e) {
            throw new RuntimeException(e);
        }
    }


    public static void carteraBruta(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void carteraDias(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangIni, int rangFin,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "dias_de_mora";

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, rangIni, rangFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void calificacion(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "calificacion";

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, calificacion, calificacion);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void reEstCapital(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            String reEstCapital = "re_est";

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, reEstCapital, 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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


    public static void reEstCapital(String okCarteraFile, int diasMoradesde, int diasMoraHasta, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String reEstCapital = "re_est";
            String diasDeMora = "dias_de_mora";
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, reEstCapital, 1, 1, diasDeMora, diasMoradesde, diasMoraHasta);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void reEstNCreditos(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void nCreditosVigentes(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void reestructuradosCapitalLc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion,  String tempFile, List <String> machSheets) throws IOException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, "calificacion", calificacion, calificacion, "re_est", 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void reestructuradosCapitalOfcPlazosProm(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, "re_est", 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularPromedioPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void reestructuradosCapitalOfcPlazosMin(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, "re_est", 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularMinimoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void reestructuradosCapitalOfcPlazosMax(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, "re_est", 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularMaximoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void mora1raCuotaMontoOfc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, campoFiltrar, valorInicio, valorFin, "cuota_desde_mora", 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void mora1raCuotaCantOfc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, campoFiltrar, valorInicio, valorFin, "cuota_desde_mora", 1, 1);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void provisiones(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "prov_cap");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void clientes(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void colocacion(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNND(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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


    public static void credPromColocacion(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNND(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularPromedioPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void colocacionPercentil(String okCarteraFile, String masterFile, String azureFile, int percent, String mesAnoCorte, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNND(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), percent);
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void cartera(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangIni, int rangFin,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "dias_de_mora";

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, rangIni, rangFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void carteraBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "calificacion";

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, calificacion, calificacion);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void pzoPromOfc(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNND(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void plazoPercentil(String okCarteraFile, String masterFile, String azureFile, int percent, String mesAnoCorte, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNND(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), percent);
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void carteraMoraCronograma(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangIni, int rangFin,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "MORA_MAXIMA_MES";

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, rangIni, rangFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void clientesMoraCronograma(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangIni, int rangFin,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");

            String campoFiltrar = "MORA_MAXIMA_MES";

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, rangIni, rangFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void modificadosCapital(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "monitoreo_modificado";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 1; // Reemplaza con el valor de fin del rango


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void modificados(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangIni, int rangFin,  String tempFile, List <String> machSheets) throws IOException, ParseException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 30; // Reemplaza con el valor de fin del rango


            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, campoFiltrar, valorInicio, valorFin, "monitoreo_modificado", rangIni, rangFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void modificadosNCreditos(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "monitoreo_modificado";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 1; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void modificadosCapital(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion,  String tempFile, List <String> machSheets) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            String campoFiltrar = "monitoreo_modificado";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 1; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, "calificacion", calificacion, calificacion, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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


    public static void modificadoCapitalOfcPlazosProm(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "monitoreo_modificado";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 1; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void modificadoCapitalOfcPlazosMin(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {


        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "monitoreo_modificado";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 1; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularMinimoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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

    public static void modificadoCapitalOfcPlazosMax(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,  String tempFile, List <String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");

            String campoFiltrar = "monitoreo_modificado";
            int valorInicio = 1; // Reemplaza con el valor de inicio del rango
            int valorFin = 1; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

            workbook.close();
            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularMaximoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null"){
                    errorMessage("Hay un null en: " + entryOkCartera.getKey());
                }

                if (datosMasterFile != null) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            //System.out.println("ENTRA AL ANALISIS ENTRE OK Y MAESTRO_for " + entry.getKey());

                            /*------------------------------------------------------------*/
                            if (entry.getKey() == "null" || entry.getValue() == "null") {
                                errorMessage("Los datos del Maestro contienen null");
                            }

                            //System.out.println("SI ESTA ENTRANDO A LA COMPARACIÓN DE DATOS ENTRE MAESTRO Y OKCARTERA");
                            if (entryOkCartera.getKey().contains(entry.getKey()) && !entryOkCartera.getKey().equals("0") && !entry.getKey().equals("0")) {

                                System.out.println("CODIGO ENCONTRADO");

                                if (!entryOkCartera.getValue().equals(entry.getValue()) || entryOkCartera.getValue() != entry.getValue() ||
                                        !entryOkCartera.getValue().contains(entry.getValue())) {

                                    String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(error);
                                    errores.add(error);

                                } else {

                                    String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                    System.out.println(coincidencia);
                                    coincidencias.add(coincidencia);

                                }
                            } else {
                                //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            /*-------------------------------------------------------------------*/
                        }
                    }
                }else {
                    String error = "La información está incompleta, no es posible completar el análisis. " +
                            "\n Por favor complete en caso de ser necesario";
                    errorMessage(error);
                    errores.add(error);
                    break;
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
