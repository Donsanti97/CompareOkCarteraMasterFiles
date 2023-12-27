package org.utils.configuration;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.utils.FunctionsApachePoi;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.utils.MethotsAzureMasterFiles.*;
import static org.utils.MethotsAzureMasterFiles.runtime;

public class GetMasterAnalasis extends FunctionsApachePoi {

    List<String> getReformatSheets(String azureFile, String masterFile, String hoja, String fechaCorte){
        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");

        List<Map<String, String>> valoresEncabezados2;
        List<Map<String, String>> valoresEncabezados1;
        List<Map<String, String>> mapList = new ArrayList<>();
        List<String> dataList;
        List<String> sht2 = new ArrayList<>();

        try {

            if (azureFile == null || masterFile == null) {
                System.out.println("No se seleccionó ningún archivo.");
            } else {
                System.out.println("Archivos válidos, el análisis comenzará en breve...");
            }

            List<String> nameSheets1 = new ArrayList<>();
            List<String> nameSheets2 = new ArrayList<>();
            assert azureFile != null;
            assert masterFile != null;
            Workbook workbook = WorkbookFactory.create(new File(azureFile));
            Workbook workbook2 = WorkbookFactory.create(new File(masterFile));
            Sheet sheet1 = null;
            Sheet sheet2 = null;

            int indexF2 = 0;


            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheet1 = workbook.getSheetAt(i);
                nameSheets1.add(sheet1.getSheetName());
            }
            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                sheet2 = workbook2.getSheetAt(i);
                nameSheets2.add(sheet2.getSheetName());
            }

            String sht1 = "";
            dataList = createDualDropDownListsAndReturnSelectedValues(nameSheets1, nameSheets2);

            for (String seleccion : dataList) {
                String[] elementos = seleccion.split(SPECIAL_CHAR);

                sht2.add(elementos[1]);
            }

            workbook.close();
            workbook2.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return sht2;
    }

    public static List<Map<String, String>> getInfo(String azureFile, String masterFile, String hoja, String fechaCorte) {
        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");

        List<Map<String, String>> valoresEncabezados2;
        List<Map<String, String>> valoresEncabezados1;
        List<Map<String, String>> mapList = new ArrayList<>();

        try {

            if (azureFile != null && masterFile != null) {
                System.out.println("Archivos válidos, el análisis comenzará en breve...");
            } else {
                System.out.println("No se seleccionó ningún archivo.");
            }

            List<String> nameSheets1 = new ArrayList<>();
            List<String> nameSheets2 = new ArrayList<>();
            assert azureFile != null;
            assert masterFile != null;
            Workbook workbook = WorkbookFactory.create(new File(azureFile));
            Workbook workbook2 = WorkbookFactory.create(new File(masterFile));
            Sheet sheet1 = null;
            Sheet sheet2 = null;

            int indexF2 = 0;


            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheet1 = workbook.getSheetAt(i);
                nameSheets1.add(sheet1.getSheetName());
            }
            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                sheet2 = workbook2.getSheetAt(i);
                nameSheets2.add(sheet2.getSheetName());
            }

            String sht1 = "";
            List<String> sht2 = new ArrayList<>();
            List<String> dataList = createDualDropDownListsAndReturnSelectedValues(nameSheets1, nameSheets2);
            List<String> encabezados1 = null;
            List<String> encabezados2 = null;
            String encabezado = "";

            for (String seleccion : dataList) {
                String[] elementos = seleccion.split(SPECIAL_CHAR);

                sht1 = elementos[0];
                sht2.add(elementos[1]);
                System.out.println("ELEMENTOS SELECCIONADOS: " + sht1 + ", " + sht2);


            }



            sheet1 = workbook.getSheet(sht1);
            encabezados1 = getHeadersN(sheet1);

            JOptionPane.showMessageDialog(null, "Del siguiente menú escoja el primer encabezado ubicado en las hojas del archivo Maestro");
            encabezado = mostrarMenu(encabezados1);

            for (int i = 0; i < sht2.size(); i++) {
                sheet2 = workbook2.getSheetAt(i);
                encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezado);
            }


            JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" que será analizado");
            String codigo = mostrarMenu(encabezados2);
            JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda a la \"Fecha de corte\" que será analizada");
            String fechaCorteMF = mostrarMenu(encabezados2);

            for (String select : sht2) {
                if (select.equals(hoja)) {
                    System.out.println("HOJA EN METODO " + select);
                    if (!fechaCorte.equals(fechaCorteMF)) {
                        errorMessage("Por favor verifique que los encabezados correspondientes a las fechas" +
                                "\n tengan un formato tipo FECHA idéntica a " + fechaCorte);

                        errorMessage("No es posible completar el análisis de la hoja [" + select +
                                "]\n el formato de fecha no es el correcto");
                    } else {
                        valoresEncabezados2 = obtenerValoresPorFilas(workbook, workbook2, sht1, select, codigo, fechaCorteMF);
                        /*mapList = createMapList(valoresEncabezados2, codigo, fechaCorteMF);
                        for (Map<String, String> map : mapList) {
                            System.out.println("Analizando valores... ");
                            for (Map.Entry<String, String> entry : map.entrySet()) {
                                System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
                            }
                        }*/
                        mapList = getInfoMaster(valoresEncabezados2, codigo, fechaCorte);
                    }
                }
            }


            System.out.println("---------------------------------------------------------------------------------------");
            System.setProperty("org.apache.poi.ooxml.strict", "true");

            System.out.println("Análisis completado...");
            workbook.close();
            workbook2.close();
            runtime();
            waitSeconds(2);


        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return mapList;
    }

    public static List<Map<String, String>> getInfoMaster(List<Map<String, String>> data, String codigo, String fechaCorte){
        List<Map<String, String>> mapList = createMapList(data, codigo, fechaCorte);
        for (Map<String, String> map : mapList) {
            System.out.println("Analizando valores... ");
            for (Map.Entry<String, String> entry : map.entrySet()) {
                System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
            }
        }
        return mapList;
    }




}
