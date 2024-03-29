package org.utils.configuration;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.utils.FunctionsApachePoi.*;
import static org.utils.MethotsAzureMasterFiles.*;

public class GetMasterAnalisis {

    public static List<String> errores = new ArrayList<>();
    public  static List<String> coincidencias = new ArrayList<>();
    /**
     * Función: este método retorna una lista que corresponde al match entre las hojas de los archivos maestro y azure
     *
     **/
    public static List<String> machSheets(String azureFile, String masterFile){
        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");

        List<String> dataList;

        try {

            System.out.println(azureFile != null && masterFile != null ? "Archivos válidos, el análisis comenzará en breve..." : "No se seleccionó ningún archivo.");

            List<String> nameSheets1 = new ArrayList<>();
            List<String> nameSheets2 = new ArrayList<>();
            assert azureFile != null;
            assert masterFile != null;
            Workbook workbook = createWorkbook(azureFile);
            Workbook workbook2 = createWorkbook(masterFile);
            Sheet sheet1;
            Sheet sheet2;


            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheet1 = workbook.getSheetAt(i);
                nameSheets1.add(sheet1.getSheetName());
            }
            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                sheet2 = workbook2.getSheetAt(i);
                nameSheets2.add(sheet2.getSheetName());
            }


            dataList = createDualDropDownListsAndReturnSelectedValues(nameSheets1, nameSheets2);

            System.setProperty("org.apache.poi.ooxml.strict", "true");

            System.out.println("Análisis completado...");
            workbook.close();
            workbook2.close();
            runtime();
            waitSeconds(2);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return dataList;
    }


    /**
     * Función: Este método retorna una lista de keys y valores que corresponden específicamente a la hoja que está siendo analizada en su momento
     * **/
    public static List<Map<String, String>> getSheetInformation(String azureFile, String masterFile, String hoja){
        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        List<Map<String, String>> valoresEncabezados2;
        List<Map<String, String>> mapList = new ArrayList<>();
        List<String> sheets = new ArrayList<>();
        try{
            Workbook workbook = createWorkbook(azureFile);
            Workbook workbook2 = createWorkbook(masterFile);
            Sheet sheet1;
            Sheet sheet2;
            Sheet sheetName;
            //String sht1 = "";
            //List<String> sht2 = new ArrayList<>();
            List<String> encabezados1;
            List<String> encabezados2;
            //String encabezado;
            int i = 0;
            String message;

            /*for (String seleccion : dataList) {
                String[] elementos = seleccion.split(SPECIAL_CHAR);
                sht1 = elementos[0];
                sht2.add(elementos[1]);
                //System.out.println("ELEMENTOS SELECCIONADOS: " + sht1 + ", " + sht2);
            }*/
            String azureSheet = workbook.getSheetName(0);

            for (int j = 0; j < workbook2.getNumberOfSheets(); j++) {
                sheets.add(workbook2.getSheetName(j));
            }

            for (String sheet : sheets) {

                if (sheet.equals(hoja)) {

                    System.out.println();
                    System.out.println("SE ESTA ANALIZANDO LA HOJA: " + hoja);

                    sheet1 = workbook.getSheet(azureSheet);
                    sheet2 = workbook2.getSheet(sheet);

                    encabezados1 = getHeadersN(sheet1);

                    encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezados1);
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" que será analizado");
                    String codigo = mostrarMenu(encabezados2);
                    while (codigo == null || codigo == "Ninguno"){
                        errorMessage("No fue seleccionado el código. Por favor siga la instrucción");
                        JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" que será analizado");
                        codigo = mostrarMenu(encabezados2);
                    }

                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Maestro de los valores que desea comparar");
                    String fechaCorteMF = mostrarMenu(encabezados2);
                    if (fechaCorteMF == null || fechaCorteMF == "Ninguno" || fechaCorteMF == "0") {
                        String yesNoAnswer = showYesNoDialog("Ha seleccionado una opción no válida.. " +
                                "\nEl encabezado con los valores que desea comparar se encuentra dentro de la lista anterior?");
                        if (yesNoAnswer.equals("SI")){
                            encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezados1);

                            JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Maestro que será analizado");
                            fechaCorteMF = mostrarMenu(encabezados2);
                            while (fechaCorteMF == null || fechaCorteMF.equals("Nunguno")) {
                                errorMessage("No fue seleccionado el encabezado. Por favor siga la instrucción");
                                JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Maestro que será analizado");
                                fechaCorteMF = mostrarMenu(encabezados2);
                            }
                            valoresEncabezados2 = obtenerValoresPorFilas(workbook, workbook2, azureSheet, sheet, codigo, fechaCorteMF, encabezados1);
                            if (valoresEncabezados2 != null){
                                System.out.println(" SI ESTÁ ENTRANDO A LLENAR EL MAPLIST DE LOS DATOS MAESTROS");
                                mapList = createMapList(valoresEncabezados2, codigo, fechaCorteMF);
                            }else {
                                message = "No es posible analizar los valores ya que los campos están incompletos." +
                                        "\n Por favor verifique que la cantidad de campos sea equivalente a la de valores. Hoja: [" + sheet + "]";
                                errorMessage(message);
                                workbook.close();
                                workbook2.close();
                                errores.add(message);
                                return null;
                            }
                        }else {
                            message = "No es posible completar el análisis de la hoja [" + hoja +
                                    "]\n el formato de fecha no es el correcto";
                            errorMessage("Por favor verifique que los encabezados a analizar existan, " +
                                    "o verifique que su archivo Excel no tenga errores en sus celdas");

                            errorMessage(message);
                            workbook.close();
                            workbook2.close();
                            errores.add(message);
                            return null;
                        }
                    } else {
                        valoresEncabezados2 = obtenerValoresPorFilas(workbook, workbook2, azureSheet, sheet, codigo, fechaCorteMF, encabezados1);
                        if (valoresEncabezados2 != null){
                            //System.out.println(" SI ESTÁ ENTRANDO A LLENAR EL MAPLIST DE LOS DATOS MAESTROS");
                            mapList = createMapList(valoresEncabezados2, codigo, fechaCorteMF);
                        }else {
                            message = "No es posible analizar los valores ya que los campos están incompletos." +
                                    "\n Por favor verifique que la cantidad de campos sea equivalente a la de valores. Hoja: [" + sheet + "]";
                            errorMessage(message);
                            workbook.close();
                            workbook2.close();
                            errores.add(message);
                            return null;
                        }
                    }
                }
            }
            System.out.println("---------------------------------------------------------------------------------------");
            System.setProperty("org.apache.poi.ooxml.strict", "true");

            System.out.println();
            workbook.close();
            workbook2.close();
            runtime();
            waitSeconds(2);
            System.out.println("Análisis completado...");

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return mapList;
    }

    public static String parsearFecha(String fechaString) {
        SimpleDateFormat formatoEntrada = new SimpleDateFormat("dd-MMM-yy", new Locale("es", "ES"));
        SimpleDateFormat formatoSalida = new SimpleDateFormat("dd/MM/yyyy");

        try {
            if (fechaString == null || fechaString.equals("Ninguno") || fechaString.equals("0")) {
                System.err.println("Fecha no encontrada");
                return null;
            } else {
                Date fecha = formatoEntrada.parse(fechaString);
                return formatoSalida.format(fecha);
            }
        } catch (ParseException e) {
            e.printStackTrace();
        }

        return null; // o manejar el error de alguna manera
    }

    /*public static void playSystemSound() {
        try {
            // Obtener el clip de sonido del sistema
            Clip clip = AudioSystem.getClip();
            // Obtener el archivo de sonido del sistema para terminar la tarea
            clip.open(AudioSystem.getAudioInputStream(GetMasterAnalisis.class.getResourceAsStream("/SystemSounds/Windows/Windows Shutdown.wav")));
            // Reproducir el sonido
            clip.start();
            // Esperar hasta que el sonido termine de reproducirse
            clip.addLineListener(new LineListener() {
                @Override
                public void update(LineEvent event) {
                    if (event.getType() == LineEvent.Type.STOP) {
                        clip.close(); // Cerrar el clip después de que termina la reproducción del sonido
                    }
                }
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/

    public static String showYesNoDialog(String message) {
        JFrame frame = new JFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // El mensaje que se mostrará en el cuadro de diálogo
        String[] options = {"SI", "NO"};

        // Mostrar el cuadro de diálogo con los botones "Sí" y "No"
        int choice = JOptionPane.showOptionDialog(
                frame,
                message,
                "Confirmación",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                options,
                options[0]
        );

        frame.dispose();

        // Retorna la opción seleccionada como String
        return (choice == JOptionPane.YES_OPTION) ? "SI" : "NO";
    }


}
