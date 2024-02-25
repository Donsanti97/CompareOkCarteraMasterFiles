package org.utils.configuration.historicoCarteraSegMonto_ColocPorOF;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
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


public class HistoricoCarteraSegMonto_ColocPorOF {
    //110 hojas

    public static boolean isEqual(String azureFile) {
        boolean isEqual = false;
        File aFile = new File(azureFile);
        if (aFile.getName().toLowerCase().contains("monto_col_oficina")) {
            isEqual = true;
        }
        return isEqual;
    }

    private static String menu(List<String> opciones) {

        JFrame frame = new JFrame("Menú de Opciones");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JComboBox<String> comboBox = new JComboBox<>(opciones.toArray(new String[0]));
        comboBox.setSelectedIndex(0);

        JButton button = new JButton("Seleccionar");

        ActionListener actionListener = new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                frame.dispose(); // Cerrar la ventana después de seleccionar una opción
            }
        };
        button.addActionListener(actionListener);

        JPanel panel = new JPanel();
        panel.add(comboBox);
        panel.add(button);

        frame.add(panel);
        frame.setSize(300, 100);
        frame.setVisible(true);

        while (frame.isVisible()) {
            // Esperar hasta que la ventana se cierre
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return comboBox.getSelectedItem().toString();
    }

    public static void configuracion(String masterFile) {

        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        while (!isEqual(azureFile)) {
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
        while (azureFile == null || okCartera == null || mesAnoCorte == null || fechaCorte == null) {
            errorMessage("Alguno de los items requeridos anteriormente no fue seleccionado." +
                    "\n Por favor seleccione nuevamente los items requeridos.");
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
            azureFile = getDocument();
            while (!isEqual(azureFile)) {
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
            waitSeconds(3);
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitSeconds(5);

            List<String> machSheets = machSheets(azureFile, masterFile);


            JOptionPane.showMessageDialog(null, "Para los análisis de algunas de las hojas a continuación es necesario que" +
                    "\n Digite a continuación un tipo de calificación entre [B] y [E]");
            List<String> opciones = Arrays.asList("B", "C", "D", "E");
            String calificacion = menu(opciones);

            nuevosOficinas(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas", tempFile, machSheets);

            nuevosOficinasMay30(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas > 30", tempFile, machSheets);

            nuevosOficinasBE(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas_B_E", calificacion, tempFile, machSheets);

            renovadoOficinas(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas", tempFile, machSheets);

            renovadoOficinasMay30(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas_>30", tempFile, machSheets);

            renovadoOficinasBE(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas_B_E", calificacion, tempFile, machSheets);

            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc '0-0.5 M", 0, 5, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 0.5-1 M", 5, 10, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 1-2 M", 10, 20, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 2-3 M", 20, 30, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 3-4 M", 30, 40, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 4-5 M", 40, 50, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 5-10 M", 50, 100, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 10-15 M", 100, 150, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 15-20 M", 150, 200, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 20-25 M", 200, 250, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 25-50 M", 250, 500, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 50-100 M", 500, 1000, tempFile, machSheets);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc > 100 M", 1000, 10000, tempFile, machSheets);

            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol '0-0.5 >30", 0, 5, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 0.5-1 > 30", 5, 10, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 1-2M >30", 10, 20, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 2-3M >30", 20, 30, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 3-4M >30", 30, 40, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 4-5M >30", 40, 50, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 5-10M >30", 50, 100, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 10-15 >30", 100, 150, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 15-20 >30", 150, 200, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 20-25 >30", 200, 250, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 25-50 >30", 250, 500, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 50-100 >30", 500, 1000, tempFile, machSheets);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol > 100 >30", 1000, 10000, tempFile, machSheets);

            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol '0-0.5 B_E", 0, 5, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 0.5-1 B_E", 5, 10, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 1-2 B_E", 10, 20, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 2-3 B_E", 20, 30, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 3-4 B_E", 30, 40, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 4-5 B_E", 40, 50, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 5-10 B_E", 50, 100, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 10-15 B_E", 100, 150, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 15-20 B_E", 150, 200, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 20-25 B_E", 200, 250, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 25-50 B_E", 250, 500, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 50-100 B_E", 500, 1000, calificacion, tempFile, machSheets);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol > 100 B_E", 1000, 10000, calificacion, tempFile, machSheets);

            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_'0-0.5 M", 0, 5, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_0.5-1 M", 5, 10, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_1-2M M", 10, 20, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_2-3M M", 20, 30, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_3-4M M", 30, 40, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_4-5M M", 40, 50, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_5-10M M", 50, 100, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_10-15 M", 100, 150, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_15-20 M", 150, 200, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_20-25 M", 200, 250, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_25-50 M", 250, 500, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_50-100 M", 500, 1000, tempFile, machSheets);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_> 100 M", 1000, 10000, tempFile, machSheets);

            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_'0-0.5 M", 0, 5, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_0.5-1 M", 5, 10, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_1-2M M", 10, 20, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_2-3M M", 20, 30, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_3-4M M", 30, 40, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_4-5M M", 40, 50, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_5-10M M", 50, 100, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_10-15 M", 100, 150, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_15-20 M", 150, 200, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_20-25 M", 200, 250, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_25-50 M", 250, 500, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_50-100 M", 500, 1000, tempFile, machSheets);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_> 100 M", 1000, 10000, tempFile, machSheets);

            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_'0-0.5 M", 0, 5, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_0.5-1 M", 5, 10, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_1-2M M", 10, 20, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_2-3M M", 20, 30, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_3-4M M", 30, 40, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_4-5M M", 40, 50, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_5-10M M", 50, 100, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_10-15 M", 100, 150, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_15-20 M", 150, 200, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_20-25 M", 200, 250, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_25-50 M", 250, 500, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_50-100 M", 500, 1000, tempFile, machSheets);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_> 100 M", 1000, 10000, tempFile, machSheets);

            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_'0-0.5 M", 0, 5, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_0.5-1 M", 5, 10, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_1-2M M", 10, 20, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_2-3M M", 20, 30, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_3-4M M", 30, 40, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_4-5M M", 40, 50, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_5-10M M", 50, 100, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_10-15 M", 100, 150, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_15-20 M", 150, 200, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_20-25 M", 200, 250, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_25-50 M", 250, 500, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_50-100 M", 500, 1000, mesAnoCorte, tempFile, machSheets);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_> 100 M", 1000, 10000, mesAnoCorte, tempFile, machSheets);


            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            logWinsToFile(masterFile, coincidencias);
            logErrorsToFile(masterFile, errores);

            deleteTempFile(tempFile);
        } catch (HeadlessException | IOException | ParseException | InterruptedException e) {
            throw new RuntimeException(e);
        }
    }


    public static void nuevosOficinas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile, List<String> machSheets) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                List<Map<String, Object>> values = getHeaderFilterValuesNS(sheet, headers, "tipo_cliente", "Nuevo", "Nuevo");

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, values, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }


    public static void nuevosOficinasMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);

                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);

                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 31, 5000);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                // Crear una nueva hoja Excel con los datos filtrados
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);


                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void nuevosOficinasBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);

                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Nuevo"; //
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSS(sheet, headers, campoFiltrar, valorInicio, valorFin, "calificacion", calificacion, calificacion);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);


                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void renovadoOficinas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);

                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Renovado"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Renovado"; //
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, valorInicio, valorFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                // Crear una nueva hoja Excel con los datos filtrados
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);


                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void renovadoOficinasMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);

                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Renovado"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Renovado"; //
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 31, 5000);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);


                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void renovadoOficinasBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);

                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Renovado"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Renovado"; //
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSS(sheet, headers, campoFiltrar, valorInicio, valorFin, "calificacion", calificacion, calificacion);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);


                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void oficinasMontoColoc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);

                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 100000; // Reemplaza con el valor de inicio del rango
                int valorFin = valorFinal * 100000; //
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void oficinasMontoColocMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
                int valorFin = valorFinal * 1000000; //

                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 31, 5000);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void oficinasMontoColocBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String calificacion, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 1000000;
                int valorFin = valorFinal * 1000000;

                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNSN(sheet, headers, "calificacion", calificacion, calificacion, campoFiltrar, valorInicio, valorFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void reestOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
                int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, "re_est", 1, 1, campoFiltrar, valorInicio, valorFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void clientesOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
                int valorFin = valorFinal * 1000000; //

                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void operacionesOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile, List<String> machSheets) throws IOException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
                int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

    public static void colocacionOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String mesAnoCorte, String tempFile, List<String> machSheets) throws IOException, ParseException, InterruptedException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

            if (datosMasterFile == null) {
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            } else {

                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "valor_desem";
                int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
                int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango
                String fechaInicio = "01/" + mesAnoCorte;
                String fechafin = "31/" + mesAnoCorte;
                Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
                Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);

                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNND(sheet, headers, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

                workbook.close();
                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, machSheets, hoja, fechaCorte);

                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {

                    if (entryOkCartera.getKey() == "null" || entryOkCartera.getValue() == "null") {
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

                                    if (entry.getValue() == entryOkCartera.getValue() || entry.getValue().contains(entryOkCartera.getValue())) {
                                        String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(coincidencia);
                                        coincidencias.add(coincidencia);
                                    } else {
                                        String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                        System.out.println(error);
                                        errores.add(error);
                                    }
                                } else {
                                    //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                                }
                                /*-------------------------------------------------------------------*/
                            }
                        }
                    } else {
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
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");

    }

}
