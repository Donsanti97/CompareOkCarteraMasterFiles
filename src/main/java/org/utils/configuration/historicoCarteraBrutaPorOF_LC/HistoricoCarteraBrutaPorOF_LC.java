package org.utils.configuration.historicoCarteraBrutaPorOF_LC;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.utils.FunctionsApachePoi.*;
import static org.utils.MethotsAzureMasterFiles.*;
import static org.utils.configuration.GetMasterAnalisis.*;


public class HistoricoCarteraBrutaPorOF_LC {
    //197 hojas

    public static boolean isEqual(String azureFile){
        boolean isEqual = false;
        File aFile = new File(azureFile);
        if (aFile.getName().toLowerCase().contains("cartera oficina_lc")){
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
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = showDateChooser();

        while (azureFile == null || okCartera == null || fechaCorte == null){
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
            JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
            fechaCorte = showDateChooser();
        }
        JOptionPane.showMessageDialog(null, "A continuación se creará un archivo temporal " +
                "\n Se recomienda seleccionar la carpeta \"Documentos\" para esta función...");
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";



        try {
            waitSeconds(3);
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(3);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitSeconds(5);

            //List<String> machSheets = machSheets(azureFile, masterFile);
            

            carteraTotal(okCartera, masterFile, azureFile, "CARTERA TOTAL", tempFile);


            hCodigoHoja(okCartera, masterFile, azureFile, "1001", 1001, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1002", 1002, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1003", 1003, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1004", 1004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1005", 1005, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1006", 1006, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1007", 1007, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1008", 1008, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1009", 1009, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1010", 1010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1011", 1011, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1017", 1017, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1025", 1025, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1026", 1026, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1028", 1028, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1029", 1028, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1030", 1030, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1032", 1032, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1034", 1034, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1035", 1035, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1036", 1036, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1037", 1037, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1038", 1038, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1039", 1039, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1041", 1041, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1042", 1042, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1044", 1044, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1047", 1047, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1049", 1049, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1050", 1050, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1051", 1051, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1053", 1053, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1055", 1055, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1056", 1056, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1087", 1087, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1088", 1088, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1089", 1089, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1093", 1093, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1094", 1094, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1095", 1095, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1096", 1096, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1097", 1097, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1202", 1202, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1203", 1203, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1221", 1221, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1233", 1233, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1234", 1234, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1236", 1236, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1237", 1237, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1238", 1238, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1701", 1701, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1703", 1703, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1704", 1704, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1705", 1705, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1706", 1706, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1801", 1801, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1819", 1819, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1901", 1901, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1902", 1902, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1904", 1904, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1905", 1905, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "1906", 1906, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2004", 2004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2006", 2006, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2007", 2007, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2008", 2008, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2009", 2009, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2010", 2010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2015", 2015, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2019", 2019, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2020", 2020, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2021", 2021, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2022", 2022, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2023", 2023, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "2024", 2024, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3001", 3001, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3004", 3004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3006", 3006, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3007", 3007, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3008", 3008, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3009", 3009, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3010", 3010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3011", 3011, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3012", 3012, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3015", 3015, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3016", 3016, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3017", 3017, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3018", 3018, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3019", 3019, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3021", 3021, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3022", 3022, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3023", 3023, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3024", 3024, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3025", 3025, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3027", 3027, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3028", 3028, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3029", 3029, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3030", 3030, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3031", 3031, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "3032", 3032, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4001", 4001, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4003", 4003, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4004", 4004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4005", 4005, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4006", 4006, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4007", 4007, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4008", 4008, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4009", 4009, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4010", 4010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4011", 4011, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4010", 4010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4013", 4013, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4014", 4014, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4015", 4015, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4016", 4016, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4017", 4017, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "4018", 4018, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5001", 5001, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5004", 5004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5006", 5006, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5007", 5007, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5008", 5008, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5009", 5009, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5010", 5010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "5011", 5011, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "6003", 6003, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "6004", 6004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "6005", 6005, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7001", 7001, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7003", 7003, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7004", 7004, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7005", 7005, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7006", 7006, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7007", 7007, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7008", 7008, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7009", 7009, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7010", 7010, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7011", 7011, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7012", 7012, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7013", 7013, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7015", 7015, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7016", 7016, tempFile);
            hCodigoHoja(okCartera, masterFile, azureFile, "7017", 7017, tempFile);

            carteraTotalMay30(okCartera, masterFile, azureFile, "CARTERA TOTAL > 30", 613, 7010, tempFile);

            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1001 > 30", 1001, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1002 > 30", 1002, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1003 > 30", 1003, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1004 > 30", 1004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1005 > 30", 1005, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1006 > 30", 1006, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1007 > 30", 1007, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1008 > 30", 1008, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1009 > 30", 1009, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1010 > 30", 1010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1011 > 30", 1011, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1017 > 30", 1017, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1025 > 30", 1025, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1026 > 30", 1026, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1028 > 30", 1028, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1029 > 30", 1028, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1030 > 30", 1030, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1032 > 30", 1032, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1034 > 30", 1034, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1035 > 30", 1035, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1036 > 30", 1036, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1037 > 30", 1037, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1038 > 30", 1038, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1039 > 30", 1039, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1041 > 30", 1041, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1042 > 30", 1042, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1044 > 30", 1044, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1047 > 30", 1047, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1049 > 30", 1049, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1050 > 30", 1050, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1051 > 30", 1051, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1053 > 30", 1053, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1055 > 30", 1055, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1056 > 30", 1056, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1087 > 30", 1087, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1088 > 30", 1088, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1089 > 30", 1089, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1093 > 30", 1093, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1094 > 30", 1094, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1095 > 30", 1095, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1096 > 30", 1096, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1097 > 30", 1097, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1202 > 30", 1202, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1203 > 30", 1203, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1221 > 30", 1221, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1233 > 30", 1233, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1234 > 30", 1234, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1236 > 30", 1236, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1237 > 30", 1237, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1238 > 30", 1238, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1701 > 30", 1701, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1703 > 30", 1703, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1704 > 30", 1704, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1705 > 30", 1705, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1706 > 30", 1706, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1801 > 30", 1801, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1819 > 30", 1819, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1901 > 30", 1901, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1902 > 30", 1902, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1904 > 30", 1904, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1905 > 30", 1905, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "1906 > 30", 1906, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2004 > 30", 2004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2006 > 30", 2006, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2007 > 30", 2007, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2008 > 30", 2008, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2009 > 30", 2009, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2010 > 30", 2010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2015 > 30", 2015, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2019 > 30", 2019, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2020 > 30", 2020, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2021 > 30", 2021, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2022 > 30", 2022, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2023 > 30", 2023, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "2024 > 30", 2024, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3001 > 30", 3001, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3004 > 30", 3004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3006 > 30", 3006, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3007 > 30", 3007, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3008 > 30", 3008, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3009 > 30", 3009, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3010 > 30", 3010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3011 > 30", 3011, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3012 > 30", 3012, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3015 > 30", 3015, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3016 > 30", 3016, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3017 > 30", 3017, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3018 > 30", 3018, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3019 > 30", 3019, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3021 > 30", 3021, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3022 > 30", 3022, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3023 > 30", 3023, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3024 > 30", 3024, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3025 > 30", 3025, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3027 > 30", 3027, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3028 > 30", 3028, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3029 > 30", 3029, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3030 > 30", 3030, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3031 > 30", 3031, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "3032 > 30", 3032, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4001 > 30", 4001, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4003 > 30", 4003, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4004 > 30", 4004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4005 > 30", 4005, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4006 > 30", 4006, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4007 > 30", 4007, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4008 > 30", 4008, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4009 > 30", 4009, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4010 > 30", 4010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4011 > 30", 4011, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4010 > 30", 4010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4013 > 30", 4013, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4014 > 30", 4014, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4015 > 30", 4015, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4016 > 30", 4016, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4017 > 30", 4017, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "4018 > 30", 4018, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5001 > 30", 5001, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5004 > 30", 5004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5006 > 30", 5006, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5007 > 30", 5007, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5008 > 30", 5008, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5009 > 30", 5009, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5010 > 30", 5010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "5011 > 30", 5011, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "6003 > 30", 6003, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "6004 > 30", 6004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "6005 > 30", 6005, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7001 > 30", 7001, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7003 > 30", 7003, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7004 > 30", 7004, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7005 > 30", 7005, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7006 > 30", 7006, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7007 > 30", 7007, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7008 > 30", 7008, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7009 > 30", 7009, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7010 > 30", 7010, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7011 > 30", 7011, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7012 > 30", 7012, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7013 > 30", 7013, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7015 > 30", 7015, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7016 > 30", 7016, tempFile);
            hCodigoHojaMay30(okCartera, masterFile, azureFile, "7017 > 30", 7017, tempFile);


            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            logWinsToFile(masterFile, coincidencias);
            logErrorsToFile(masterFile, errores);

            deleteTempFile(tempFile);
        } catch (HeadlessException | IOException e) {
            throw new RuntimeException(e);
        }
    }


    public static void carteraTotal(String okCarteraFile, String masterFile, String azureFile, String hoja,  String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

            if (datosMasterFile == null){
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            }else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("linea", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "dias_de_mora";
                int valorInicio = 0; // Reemplaza con el valor de inicio del rango
                int valorFin = 5000; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);
                workbook.close();

                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

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

    public static void hCodigoHoja(String okCarteraFile, String masterFile, String azureFile, String hoja, int codigoHoja,  String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

            if (datosMasterFile == null){
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            }else {

                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("linea", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
                String campoFiltrar = "codigo_sucursal";
                int valorInicio = codigoHoja;
                int valorFin = codigoHoja;

                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNN(sheet, headers, campoFiltrar, valorInicio, valorFin);
                workbook.close();


                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

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

    public static void carteraTotalMay30(String okCarteraFile, String masterFile, String azureFile, String hoja, int codigoHojaIni, int codigoHojafin,  String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

            if (datosMasterFile == null){
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            }else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("linea", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "dias_de_mora";
                int valorInicio = 31; // Reemplaza con el valor de inicio del rango
                int valorFin = 5000; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, campoFiltrar, valorInicio, valorFin, "codigo_sucursal", codigoHojaIni, codigoHojafin);
                workbook.close();

                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

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

    public static void hCodigoHojaMay30(String okCarteraFile, String masterFile, String azureFile, String hoja, int codigoHoja,  String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            String message;

            List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

            if (datosMasterFile == null){
                message = "La información está incompleta, no es posible completar el análisis. " +
                        "\n Por favor complete en caso de ser necesario. Hoja: [" + hoja + "]";
                errorMessage(message);
                errores.add(message);
            }else {
                Workbook workbook = createWorkbook(okCarteraFile);


                IOUtils.setByteArrayMaxOverride(20000000);

                Sheet sheet = workbook.getSheetAt(0);

                List<String> headers = getHeadersN(sheet);
                List<String> camposDeseados = Arrays.asList("linea", "capital");
                System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

                String campoFiltrar = "dias_de_mora";
                int valorInicio = 31; // Reemplaza con el valor de inicio del rango
                int valorFin = 5000; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNNN(sheet, headers, campoFiltrar, valorInicio, valorFin, "codigo_sucursal", codigoHoja, codigoHoja);
                workbook.close();


                System.out.println();
                System.out.println("CREANDO ARCHIVO TEMPORAL");
                crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

                workbook = createWorkbook(tempFile);

                sheet = workbook.getSheetAt(0);

                System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
                //List<Map<String, String>> datosMasterFile = getSheetInformation(azureFile, masterFile, hoja);

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
