package org.utils.configuration;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.io.File;
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
            Workbook workbook = WorkbookFactory.create(new File(azureFile));
            Workbook workbook2 = WorkbookFactory.create(new File(masterFile));
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
    public static List<Map<String, String>> getSheetInformation(String azureFile, String masterFile, List<String> dataList, String hoja, String fechaCorte){
        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        List<Map<String, String>> valoresEncabezados2;
        List<Map<String, String>> mapList = new ArrayList<>();
        try{
            Workbook workbook = WorkbookFactory.create(new File(azureFile));
            Workbook workbook2 = WorkbookFactory.create(new File(masterFile));
            Sheet sheet1;
            Sheet sheet2;
            String sht1 = "";
            List<String> sht2 = new ArrayList<>();
            List<String> encabezados1;
            List<String> encabezados2;
            String encabezado;
            int i = 0;

            for (String seleccion : dataList) {
                String[] elementos = seleccion.split(SPECIAL_CHAR);
                sht1 = elementos[0];
                sht2.add(elementos[1]);
                //System.out.println("ELEMENTOS SELECCIONADOS: " + sht1 + ", " + sht2);
            }

            for (String sheet : sht2) {

                if (sheet.equals(hoja)) {

                    System.out.println();
                    System.out.println("SE ESTA ANALIZANDO LA HOJA: " + hoja);
                    sheet1 = workbook.getSheet(sht1);
                    encabezados1 = getHeadersN(sheet1);

                    JOptionPane.showMessageDialog(null, "Del siguiente menú escoja el primer encabezado ubicado en las hojas del archivo Maestro");
                    assert encabezados1 != null;
                    encabezado = mostrarMenu(encabezados1);
                    sheet2 = workbook2.getSheet(sheet);
                    encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezado);
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" que será analizado");
                    String codigo = mostrarMenu(encabezados2);
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda a la \"Fecha de corte\" que será analizada");
                    String fechaCorteMF = mostrarMenu(encabezados2);
                    String fecha = parsearFecha(fechaCorteMF);
                    System.out.println("Fecha formateada: " + fecha);
                    if (!fechaCorte.equals(fecha)) {
                        errorMessage("Por favor verifique que los encabezados correspondientes a las fechas" +
                                "\n tengan un formato tipo FECHA idéntica a " + fechaCorte +
                                "\n o en su defecto que aparezca en la lista");

                        errorMessage("No es posible completar el análisis de la hoja [" + hoja +
                                "]\n el formato de fecha no es el correcto");
                        workbook.close();
                        workbook2.close();
                        return null;
                    } else {
                        valoresEncabezados2 = obtenerValoresPorFilas(workbook, workbook2, sht1, sheet, codigo, fechaCorteMF);
                        if (valoresEncabezados2 != null){
                            System.out.println(" SI ESTÁ ENTRANDO A LLENAR EL MAPLIST DE LOS DATOS MAESTROS");
                            mapList = createMapList(valoresEncabezados2, codigo, fechaCorteMF);
                        }else {
                            errorMessage("No es posible analizar los valores ya que los campos están incompletos." +
                                    "\n Por favor verifique que la cantidad de campos sea equivalente a la de valores.");
                            workbook.close();
                            workbook2.close();

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
            if (fechaString == null || fechaString.equals("Ninguno")) {
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

}
