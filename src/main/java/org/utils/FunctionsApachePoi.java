package org.utils;

import com.toedter.calendar.JDateChooser;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;
import java.util.concurrent.Semaphore;
import java.util.concurrent.atomic.AtomicReference;

import static org.utils.MethotsAzureMasterFiles.obtenerValorVisibleCelda;
import static org.utils.MethotsAzureMasterFiles.runtime;


public class FunctionsApachePoi {


    private static final Logger logger = LogManager.getLogger(FunctionsApachePoi.class);

    private static final int ROWS_PER_BATCH = 3000;


    //Método para obtener los valores de encabezados generales
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        value = obtenerValorVisibleCelda(cell);
                    }
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return data;
    }*/

    /*public static void convertirExcel(String archivo) throws IOException {
        FileInputStream fis = new FileInputStream(archivo);
        Workbook workbook = WorkbookFactory.create(new File(archivo));;

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        try {
                            double valorNumerico = Double.parseDouble(cell.getStringCellValue());
                            // Si se puede convertir a número, establece el valor numérico
                            cell.setCellValue(valorNumerico);
                        } catch (NumberFormatException e) {
                            // No se pudo convertir a número, no hacemos nada
                        }
                    }
                }
            }
        }

        fis.close();

        // Guardar el archivo Excel con los valores convertidos
        FileOutputStream fos = new FileOutputStream(archivo);
        workbook.write(fos);
        fos.close();

        workbook.close();
    }*/

    //@Test
    //Método para creación de tablas dinámicas
    /*public static void tablasDinamicasApachePoi(String filePath, String codSucursal, String colValores, String funcion) {

        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            convertirExcel(filePath);

            InputStream fileInputStream = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            //Definir hoja
            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = obtenerEncabezados(filePath, sheet.getSheetName());
            int index = 0;
            int index2 = 0;
            for (int i = 0; i < headers.size(); i++) {
                String header = headers.get(i);
                if (header.contains(codSucursal)) {
                    index = i;
                    System.out.println("Index1: " + index);
                }
            }
            for (int i = 0; i < headers.size(); i++) {
                String header = headers.get(i);
                if (header.contains(colValores)) {
                    index2 = i;
                    System.out.println("Index2: " + index2);

                }
            }


            //Generar el área de los datos
            CellReference topLeft = new CellReference(sheet.getFirstRowNum(), sheet.getRow(sheet.getFirstRowNum()).getFirstCellNum());
            CellReference bottomRight = new CellReference(sheet.getLastRowNum(), sheet.getRow(sheet.getLastRowNum()).getLastCellNum() - 1);
            AreaReference source = new AreaReference(topLeft, bottomRight, sheet.getWorkbook().getSpreadsheetVersion());
            System.out.println(source);


            CellReference pivotCellReference = new CellReference(2, bottomRight.getCol() + 3);

            //Crea la tabla dinámica en la hoja de trabajo
            XSSFPivotTable pivotTable = ((XSSFSheet) sheet).createPivotTable(source, pivotCellReference);//DW12
            pivotTable.addRowLabel(index);//Agregar etiqueta de fila para el campo Modalidad (12)


            switch (funcion.toLowerCase()) {
                case "suma":
                    pivotTable.addColumnLabel(DataConsolidateFunction.SUM, index2, "Suma de " + colValores);//Agrega la columna de la que se va a hacer la suma y la etiqueta de la función suma(15)
                    break;
                case "recuento":
                    pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, index2, "Recuento de " + colValores);//Agrega la columna de la que se va a hacer la suma y la etiqueta de la función suma(15)
                    break;
                case "promedio":
                    pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, index2, "Promedio de " + colValores);//Agrega la columna de la que se va a hacer la suma y la etiqueta de la función suma(15)

            }


            //Guardar excel
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            fileInputStream.close();
            fileOut.close();


            //Se cierra excel
            workbook.close();


            System.out.println("Tabla dinámica creada");

        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
    }*/


    public static void waitSeconds(int seconds) {
        try {
            Thread.sleep((seconds * 1000L));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void waitMinutes(int minutes) {
        try {
            Thread.sleep((minutes * 10000L));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /*public static Map<String, Integer> extractPivotTableData(String filePath, String filterColumnName, String valueColumnName) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(new File(filePath));;
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println("Hoja: " + sheet.getSheetName());
        List<XSSFTable> tables = ((XSSFSheet) sheet).getTables();
        System.out.println("Tablas: " + ((XSSFSheet) sheet).getTables().get(0).toString());
        if (tables.isEmpty()) {
            throw new IllegalArgumentException("No se encontraron tablas dinámicas en la hoja de trabajo.");
        }

        XSSFTable pivotTable = tables.get(0);
        CellReference startCell = pivotTable.getStartCellReference();
        CellReference endCell = pivotTable.getEndCellReference();
        int firstRow = startCell.getRow();
        int lastRow = endCell.getRow();

        Map<String, Integer> dataMap = new HashMap<>();

        for (int rowNum = firstRow + 1; rowNum <= lastRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            Cell filterCell = row.getCell(pivotTable.findColumnIndex(filterColumnName));
            String filterValue = filterCell.getStringCellValue();
            Cell valueCell = row.getCell(pivotTable.findColumnIndex(valueColumnName));
            int sumValue = (int) valueCell.getNumericCellValue();
            dataMap.put(filterValue, sumValue);
        }

        fis.close();
        return dataMap;
    }*/

    /*public static Map<String, Integer> processExcelFile(String filePath) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(new File(filePath));;
        Sheet sheet = workbook.getSheetAt(0); // Suponiendo que estás trabajando en la primera hoja del archivo

        Map<String, Integer> resultMap = new HashMap<>();

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // Saltar la primera fila (encabezados)

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell tipoProductoCell = row.getCell(0); // Suponiendo que la columna 0 contiene el tipo de producto
            Cell costoCell = row.getCell(1); // Suponiendo que la columna 1 contiene el costo por producto

            String tipoProducto = tipoProductoCell.getStringCellValue();
            int costo = (int) costoCell.getNumericCellValue();

            // Verificar si ya existe la entrada en el Map
            if (resultMap.containsKey(tipoProducto)) {
                // Sí existe, agregar el costo al valor existente
                int sumaCosto = resultMap.get(tipoProducto) + costo;
                resultMap.put(tipoProducto, sumaCosto);
            } else {
                // Si no existe, agregar una nueva entrada en el Map
                resultMap.put(tipoProducto, costo);
            }
        }

        fis.close();
        return resultMap;
    }*/

    /*-----------------------------------------------------------------------------------------------------------------------------------------*/
    //private static final int BATCH_SIZE = 1000; // Tamaño del lote para procesar
    /*-----------------------------------------------------------------------------------------------------------------------------------------*/

    //Método para obtener los nombres de las hojas existentes en el excel
    /*public static List<String> obtenerNombresDeHojas(String excelFilePath) {
        List<String> sheetNames = new ArrayList<>();
        try {

            IOUtils.setByteArrayMaxOverride(300000000);

            XSSFWorkbook workbook = new XSSFWorkbook(new File(excelFilePath));
            runtime();
            waitSeconds(2);

            XSSFSheet sheet = workbook.getSheetAt(0);
            sheetNames.add(sheet.getSheetName());

            workbook.close();
            //fis.close();
        } catch (IOException | InvalidFormatException e) {
            logger.error("Error al procesar el archivo Excel", e);
            System.err.println("Error al procesar el archivo Excel: " + e);
        }
        return sheetNames;
    }*/


    /*public static void convertExcelToCsv(String excelFilePath, String csvFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
             BufferedWriter writer = new BufferedWriter(new FileWriter(csvFilePath))) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    writer.write(obtenerValorVisibleCelda(cell));
                    writer.write(",");
                }
                writer.newLine();
            }
        }
    }*/
/*------------------------------------------------------------------------------------------------------------------------------------------*/
/*public static List<String> obtenerEncabezados(String excelFilePath, String sheetName) {
    List<String> headers = new ArrayList<>();

    try {
        //FileInputStream fis = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(new File(excelFilePath));

        // Obtener el contenido del archivo XLSX como un XmlObject
        XmlObject xmlObject = XmlObject.Factory.parse(workbook.getPackagePart().getInputStream(), new XmlOptions());

        // Crear un XMLStreamReader para leer el contenido XML
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLStreamReader reader = factory.createXMLStreamReader(xmlObject.newInputStream());

        // Iterar a través del contenido XML
        while (reader.hasNext()) {
            int event = reader.next();
            // Verificar si es un elemento de hoja y obtener el nombre de la hoja
            if (event == XMLStreamConstants.START_ELEMENT){
                if ("sheetData".equals(reader.getLocalName())) {
                    headers = readHeaderRow(reader);
                    break;
                }
            }
        }

        // Cerrar recursos
        reader.close();
        //fis.close();
        workbook.close();
    } catch (Exception e) {
        e.printStackTrace();
    }

    return headers;
}

    private static List<String> readHeaderRow(XMLStreamReader reader) throws Exception {
        List<String> headers = new ArrayList<>();
        while (reader.hasNext()) {
            int event = reader.next();
            switch (event) {
                case XMLStreamConstants.START_ELEMENT:
                    // Verificar si es un elemento de fila
                    if ("row".equals(reader.getLocalName())) {
                        // Leer la fila y agregar los encabezados
                        while (reader.hasNext()) {
                            int cellEvent = reader.next();
                            if (cellEvent == XMLStreamConstants.START_ELEMENT && "c".equals(reader.getLocalName())) {
                                headers.add(readCellValue(reader));
                            } else if (cellEvent == XMLStreamConstants.END_ELEMENT && "row".equals(reader.getLocalName())) {
                                // Fin de la fila, salir del bucle
                                break;
                            }
                        }
                        break;
                    }
                    break;
            }
        }
        return headers;
    }

    private static String readCellValue(XMLStreamReader reader) throws Exception {
        String value = "";
        while (reader.hasNext()) {
            int cellEvent = reader.next();
            if (cellEvent == XMLStreamConstants.START_ELEMENT && "v".equals(reader.getLocalName())) {
                value = reader.getElementText();
                break;
            }
        }
        return value;
    }*/
    /*---------------------------------------------------------------------------------------------------------------------------------------*/
    //Método para obtener los encabezados en las hojas
    /*public static List<String> obtenerEncabezados(String excelFilePath, String sheetName) {
        List<String> headers = new ArrayList<>();
        try {
            IOUtils.setByteArrayMaxOverride(300000000);
            //FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
            Sheet sheet = workbook.getSheet(sheetName);
            Row headerRow = sheet.getRow(0);
            String value = "";
            for (Cell cell : headerRow) {
                value = obtenerValorVisibleCelda(cell);
                headers.add(value);
            }
            workbook.close();
            //fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return headers;
    }*/

    //Método para obtener los valores de encabezados específicos
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, List<String> camposDeseados) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        value = obtenerValorVisibleCelda(cell);
                    }
                    if (camposDeseados.contains(header)) {
                        rowData.put(header, value);
                    }
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return data;
    }*/


    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, List<String> camposDeseados, String header) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String currentHeader = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            value = cell.getStringCellValue();
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            value = String.valueOf(cell.getNumericCellValue());
                        }
                    }
                    if (camposDeseados.contains(currentHeader) && currentHeader.equals(header)) {
                        rowData.put(currentHeader, value);
                    }
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return data;
    }*/

    /*---------------------------------------------------------------------------------------------------*/
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, List<String> camposDeseados, int percent) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
        try {
            convertirExcel(excelFilePath);

            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    double porcentaje = (double) percent / 100;
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            value = cell.getStringCellValue();
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            value = String.valueOf(cell.getNumericCellValue() * porcentaje);
                        }
                    }
                    if (camposDeseados.contains(header)) {
                        rowData.put(header, value);
                    }
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return data;
    }*/


    /*---------------------------------------------------------------------------------------------------*/

    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar, int valorInicio, int valorFin) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex = headers.indexOf(campoFiltrar);
            if (campoFiltrarIndex == -1) {
                System.err.println("El campo especificado para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(campoFiltrarIndex);
                double valorCelda = (cell != null && cell.getCellType() == CellType.NUMERIC) ? cell.getNumericCellValue() : 0;
                if (valorCelda >= valorInicio && valorCelda <= valorFin) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            value = obtenerValorVisibleCelda(dataCell);
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    /*-------------------------------------------------------------------------------------------------*/

   /* public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar, String valorInicio, String valorFin, List<String> datosDeseados, String tempFile) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();

        try {
             Workbook workbook = WorkbookFactory.create(new File(excelFilePath));

            Sheet sheet = workbook.getSheet(sheetName);

            List<String> headers = obtenerEncabezados(excelFilePath, sheetName*//*workbook, sheetName*//*);
            int campoFiltrarIndex = headers.indexOf(campoFiltrar);

            if (campoFiltrarIndex == -1) {
                System.err.println("El campo especificado para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();

            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(campoFiltrarIndex);

                String valorCelda = (cell != null && cell.getCellType() == CellType.STRING) ? obtenerValorVisibleCelda(cell)*//*cell.getStringCellValue()*//* : "";

                if (valorCelda.compareTo(valorInicio) >= 0 && valorCelda.compareTo(valorFin) <= 0) {
                    Map<String, String> rowData = new HashMap<>();

                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";

                        if (dataCell != null) {
                            value = obtenerValorVisibleCelda(dataCell);
                        }

                        rowData.put(header, value);
                    }

                    datosFiltrados.add(rowData);

                    *//*--------------------------------------------------------------------------------------------------------------------*//*
                    Map<String, String> datosProcesados = procesarLote(datosFiltrados, datosDeseados, tempFile);


                    datosFiltrados.add(datosProcesados);
                    *//*--------------------------------------------------------------------------------------------------------------------*//*
                }
            }
workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return datosFiltrados;
    }*/

    /*private static Map<String, String> procesarLote(List<Map<String, String>> datosFiltrados, List<String> headers, String tempFile) {
        Map<String, String> resultado;
        try {
            crearNuevaHojaExcel(tempFile, headers, datosFiltrados);
            resultado = functions.calcularSumaPorValoresUnicos(tempFile, headers.get(0), headers.get(1));
            if (datosFiltrados.size() % BATCH_SIZE == 0) {
                System.out.println("Procesando lote de filas: " + datosFiltrados.size());
                for (String header : headers) {
                    System.out.println("HEADER: " + header);
                    String dato = "";
                    for (Map<String, String> rowData : datosFiltrados) {
                        for (int i = 0; i < datosFiltrados.size(); i++) {
                            dato = rowData.get(header);
                            System.out.println(header + ": " + dato);
                        }
                        resultado.put(header, dato);
                    }
                }
                System.gc();
                waitSeconds(3);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
        return resultado;
    }*/

    /*private static List<String> obtenerEncabezados(Workbook workbook, String sheetName) {
        List<String> headers = new ArrayList<>();
        Sheet sheet = workbook.getSheet(sheetName);
        Row headerRow = sheet.getRow(0);

        for (Cell cell : headerRow) {
            //headers.add(cell.getStringCellValue());
            headers.add(obtenerValorVisibleCelda(cell));

        }

        return headers;
    }*/
    /*--------------------------------------------------------------------------------------------------*/

    //Método para obtener valores de los encabezados en un rango especifico de valores
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar, String valorInicio, String valorFin) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            //FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex = headers.indexOf(campoFiltrar);
            if (campoFiltrarIndex == -1) {
                System.err.println("El campo especificado para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(campoFiltrarIndex);
                String valorCelda = (cell != null && cell.getCellType() == CellType.STRING) ? cell.getStringCellValue() : "";
                if (valorCelda.compareTo(valorInicio) >= 0 && valorCelda.compareTo(valorFin) <= 0) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            value = obtenerValorVisibleCelda(dataCell);
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            //fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    //Método para obtener valores de dos encabezados de un rango específico de valores cada uno
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, String valorInicio1, String valorFin1, String campoFiltrar2, String valorInicio2, String valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";
                String valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.STRING) ? cell2.getStringCellValue() : "";
                if (valorCelda1.compareTo(valorInicio1) >= 0 && valorCelda1.compareTo(valorFin1) <= 0 &&
                        valorCelda2.compareTo(valorInicio2) >= 0 && valorCelda2.compareTo(valorFin2) <= 0) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(dataCell.getNumericCellValue());

                            } else if (dataCell.getCellType() == CellType.STRING && DateUtil.isCellDateFormatted(dataCell)) {
                                DataFormatter dataFormatter = new DataFormatter();
                                value = dataFormatter.formatCellValue(dataCell);
                            }
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    //Método para obtener valores de dos encabezados de un rango específico cada uno, en campos numéricos
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, double valorInicio1, double valorFin1, String campoFiltrar2, double valorInicio2, double valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                System.gc();
                waitSeconds(2);
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                double valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.NUMERIC) ? cell1.getNumericCellValue() : 0.0;
                double valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? cell2.getNumericCellValue() : 0.0;
                if (valorCelda1 >= valorInicio1 && valorCelda1 <= valorFin1 &&
                        valorCelda2 >= valorInicio2 && valorCelda2 <= valorFin2) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            value = obtenerValorVisibleCelda(dataCell);
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
                System.gc();
                waitSeconds(2);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    //Método para obtener valores de los encabezados de un rango específico cada uno, el primero rango String y el segundo rango double
    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, String valorInicio1, String valorFin1, String campoFiltrar2, int valorInicio2, int valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        List<String> headers;
        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";
                double valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? cell2.getNumericCellValue() : 0.0;
                if (valorCelda1.compareTo(valorInicio1) >= 0 && valorCelda1.compareTo(valorFin1) <= 0 &&
                        valorCelda2 >= valorInicio2 && valorCelda2 <= valorFin2) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            value = obtenerValorVisibleCelda(dataCell);
                        }

                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, int valorInicio1, int valorFin1, String campoFiltrar2, int valorInicio2, int valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        List<String> headers;
        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }

            runtime();
            waitSeconds(2);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";
                double valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? cell2.getNumericCellValue() : 0.0;

                if (valorCelda1.compareTo(String.valueOf(valorInicio1)) >= 0 && valorCelda1.compareTo(String.valueOf(valorFin1)) <= 0 &&
                        valorCelda2 >= valorInicio2 && valorCelda2 <= valorFin2) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            value = obtenerValorVisibleCelda(dataCell);
                        }

                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
                runtime();
                waitSeconds(2);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, String valorInicio1, String valorFin1, String campoFiltrar2, Date valorInicio2, Date valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }
            runtime();
            waitSeconds(2);

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);

                // Convertir celda 2 a fecha si es de tipo fecha
                Date fechaCelda2 = null;
                if (cell2 != null && cell2.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell2)) {
                    fechaCelda2 = cell2.getDateCellValue();
                }
                runtime();
                waitSeconds(2);

                // Obtener el valor de celda 1 como cadena de texto
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";

                if (fechaCelda2 != null &&
                        valorCelda1.compareTo(valorInicio1) >= 0 && valorCelda1.compareTo(valorFin1) <= 0 &&
                        fechaCelda2.compareTo(valorInicio2) >= 0 && fechaCelda2.compareTo(valorFin2) <= 0) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                if (DateUtil.isCellDateFormatted(dataCell)) {
                                    Date fecha = dataCell.getDateCellValue();
                                    value = dateFormat.format(fecha);
                                } else {
                                    value = String.valueOf(dataCell.getNumericCellValue());
                                }
                            }
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
                runtime();
                waitSeconds(2);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/

    public static void deleteTempFile(String tempFile) {
        eliminarExcel(tempFile, 5);
    }

    /*public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, int valorInicio1, int valorFin1, String campoFiltrar2, Date valorInicio2, Date valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }
            runtime();
            waitSeconds(2);

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);

                // Convertir celda 2 a fecha si es de tipo fecha
                Date fechaCelda2 = null;
                if (cell2 != null && cell2.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell2)) {
                    fechaCelda2 = cell2.getDateCellValue();
                }
                runtime();
                waitSeconds(2);

                // Obtener el valor de celda 1 como cadena de texto
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";

                if (fechaCelda2 != null &&
                        valorCelda1.compareTo(String.valueOf(valorInicio1)) >= 0 && valorCelda1.compareTo(String.valueOf(valorFin1)) <= 0 &&
                        fechaCelda2.compareTo(valorInicio2) >= 0 && fechaCelda2.compareTo(valorFin2) <= 0) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                if (DateUtil.isCellDateFormatted(dataCell)) {
                                    Date fecha = dataCell.getDateCellValue();
                                    value = dateFormat.format(fecha);
                                } else {
                                    value = String.valueOf(dataCell.getNumericCellValue());
                                }
                            }
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
                runtime();
                waitSeconds(2);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }*/


    public static Workbook createWorkbook(String filePath) {
        try {
            File file = new File(filePath);
            if (filePath.endsWith(".xls")) {
                return new HSSFWorkbook(POIFSFileSystem.create(file));
            } else {
                return new XSSFWorkbook(file);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
    //Método que crea una nueva hoja excel con información específica ya tratada en un archivo excel nuevo
    public static void crearNuevaHojaExcel(String filePath, List<String> headers, List<Map<String, String>> data) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NuevaHoja");

        // Crear la fila de encabezados en la nueva hoja
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        // Llenar la nueva hoja con los datos filtrados
        for (int i = 0; i < data.size(); i++) {
            Map<String, String> rowData = data.get(i);
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < headers.size(); j++) {
                String header = headers.get(j);
                String value = rowData.get(header);
                Cell cell = row.createCell(j);
                cell.setCellValue(value);
            }
        }


        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            System.out.println("Nueva hoja Excel creada o reemplazada en: " + filePath);
            fos.close();
            workbook.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
    }

    public static void crearNuevaHojaExcel(List<String> headers, List<Map<String, Object>> data, String filePath) throws InterruptedException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NuevaHoja");

        int count1 = 0;
        int count2 = 0;
        //int rowsPerBatch = 5000;

        // Crear la fila de encabezados en la nueva hoja
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            count1++;
        }
        if (count1 % ROWS_PER_BATCH == 0) {
            runtime();
            Thread.sleep(200);
        }

        // Llenar la nueva hoja con los datos filtrados
        for (int i = 0; i < data.size(); i++) {
            Map<String, Object> rowData = data.get(i);
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < headers.size(); j++) {
                String header = headers.get(j);
                String value = (String) rowData.get(header);
                Cell cell = row.createCell(j);
                cell.setCellValue(value);
                count2++;
            }
        }
        if (count2 % ROWS_PER_BATCH == 0) {
            runtime();
            Thread.sleep(200);
        }


        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            System.out.println("Nueva hoja Excel creada o reemplazada en: " + filePath);
            fos.close();
            workbook.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
    }

    public static void crearNuevaHojaExcel(String filePath, List<String> headers) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NuevaHoja");

        // Crear la fila de encabezados en la nueva hoja
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            System.out.println("Nueva hoja Excel creada o reemplazada en: " + filePath);
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                logger.error("Error al cerrar el libro de Excel", e);
            }
        }
    }

    //Método que elimina un archivo excel existente
    public static void eliminarExcel(String filepath, int waitSeconds) {
        File tempFile = new File(filepath);
        int seconds = waitSeconds * 1000;

        if (tempFile.exists()) {
            try {
                // Espera durante el tiempo especificado antes de eliminar el archivo
                Thread.sleep(seconds);

                if (tempFile.delete()) {
                    System.out.println("Archivo Excel temporal eliminado con éxito.");
                } else {
                    System.out.println("No se pudo eliminar el archivo Excel temporal.");
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                System.err.println("Error al esperar antes de eliminar el archivo temporal: " + e.getMessage());
            }
        } else {
            System.out.println("El archivo Excel temporal no existe.");
        }


    }

    /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
    /*public static List<String> obtenerEncabezados(Sheet sheet) {
        List<String> encabezados = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Aquí puedes especificar en qué fila esperas encontrar los encabezados
            // Por ejemplo, si están en la tercera fila (fila índice 2), puedes usar:
            if (row.getRowNum() == 0) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    encabezados.add(obtenerValorVisibleCelda(cell));
                }
                break; // Terminamos de buscar encabezados una vez que los encontramos
            }
        }

        return encabezados;
    }*/

    /*public static List<String> obtenerNombresDeHojas(String excelFilePath, int indexFrom) {
        List<String> sheetNames = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(new File(excelFilePath));;
            int numberOfSheets = workbook.getNumberOfSheets();
            for (int i = indexFrom; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                sheetNames.add(sheet.getSheetName());
            }
            workbook.close();
            fis.close();
            runtime();
            waitSeconds(2);
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return sheetNames;
    }*/

    /*public static List<String> obtenerEncabezados(Sheet sheet, int index) {
        List<String> encabezados = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Aquí puedes especificar en qué fila esperas encontrar los encabezados
            // Por ejemplo, si están en la tercera fila (fila índice 2), puedes usar:
            if (row.getRowNum() == index) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    encabezados.add(obtenerValorVisibleCelda(cell));
                }
                break; // Terminamos de buscar encabezados una vez que los encontramos
            }
        }
        waitSeconds(5);
        runtime();

        return encabezados;

    }*/

    /*public static List<String> encontrarEncabezadosSegundoArchivo(Sheet sheet, Workbook workbook2) {
        List<String> encabezadosSegundoArchivo = new ArrayList<>();

        // Busca el primer encabezado del primer archivo en la misma columna en el segundo archivo
        for (int columnIndex = 0; columnIndex < sheet.getRow(0).getLastCellNum(); columnIndex++) {
            String primerEncabezado = obtenerValorVisibleCelda(sheet.getRow(0).getCell(columnIndex));
            if (buscarEncabezadoEnColumna(primerEncabezado, columnIndex, workbook2)) {
                Sheet segundoSheet = workbook2.getSheetAt(3); // Puedes especificar el índice de la hoja del segundo archivo
                Iterator<Row> rowIterator = segundoSheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Cell cell = row.getCell(columnIndex);
                    encabezadosSegundoArchivo.add(obtenerValorVisibleCelda(cell));
                }
                break; // Terminamos de buscar encabezados en el segundo archivo
            }
        }

        return encabezadosSegundoArchivo;
    }*/

    /*private static boolean buscarEncabezadoEnColumna(String encabezado, int columnIndex, Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(3); // Puedes especificar el índice de la hoja del segundo archivo
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnIndex);
            String valor = obtenerValorVisibleCelda(cell);
            if (!valor.equals(null) || !valor.isEmpty()) {
                valor = "0";
            }
            if (encabezado.equals(valor)) {
                return true;
            }
        }
        return false;
    }*/

    /*-----------------------------------------------------------------------------------------*/
    /*public static List<String> buscarValorEnColumna(Sheet sheet, int columnaBuscada, String valorBuscado) {
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnaBuscada);
            String valorCelda = obtenerValorVisibleCelda(cell);

            if (valorBuscado.equals(valorCelda)) {
                return obtenerValoresFila(row);
            }
        }

        return null; // Valor no encontrado en la columna especificada
    }*/

   /* public static Map<String, List<String>> obtenerValoresPorEncabezado(Sheet sheet, List<String> encabezados) {
        Map<String, List<String>> valoresPorEncabezado = new HashMap<>();

        for (String encabezado : encabezados) {
            valoresPorEncabezado.put(encabezado, new ArrayList<>());
        }

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila, ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);

                if (valoresPorEncabezado.containsKey(encabezado)) {
                    valoresPorEncabezado.get(encabezado).add(valor);
                }
            }
        }

        return valoresPorEncabezado;
    }*/


    /*public static List<String> obtenerValoresFila(Row row) {
        List<String> valoresFila = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            valoresFila.add(obtenerValorVisibleCelda(cell));
        }
        runtime();
        return valoresFila;
    }*/
    /*-----------------------------------------------------------------------------------------------*/


    /*private static String obtenerValorCelda(Cell cell) {
        String valor = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    valor = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        //valor = cell.getDateCellValue().toString();
                        String formatDate = cell.getDateCellValue().toString();
                        SimpleDateFormat formatoEntrada = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                        try {
                            Date date = formatoEntrada.parse(formatDate);
                            SimpleDateFormat formatoSalida = new SimpleDateFormat("dd/MM/yyyy");
                            valor = formatoSalida.format(date);
                        } catch (ParseException e) {
                            throw new RuntimeException(e);
                        }
                    } else {
                        valor = Double.toString(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    valor = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    valor = cell.getCellFormula();
                    break;
                default:
                    break;
            }
        }
        return valor;
    }*/

    /*--------------------OTROS MÉTODOS PARA LEER Y HACER LA SUMATORIA POR VALOR---------------------------------------------------------------*/
    public static List<Map<String, String>> leerExcel(String filePath) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();

        try (Workbook workbook = createWorkbook(filePath)/*WorkbookFactory.create(new File(filePath))*/) {

            Sheet sheet = workbook.getSheetAt(0); // Supongamos que es la primera hoja

            Row headerRow = sheet.getRow(0);

            int countRows = 0;
            //int rowsPerBatch = 5000;
            int totalRows = sheet.getPhysicalNumberOfRows() - 1;
            System.out.println("LECTURA EXCEL");
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row currentRow = sheet.getRow(rowIndex);
                Map<String, String> rowMap = new HashMap<>();

                for (int columnIndex = 0; columnIndex < headerRow.getLastCellNum(); columnIndex++) {
                    Cell headerCell = headerRow.getCell(columnIndex);
                    Cell currentCell = currentRow.getCell(columnIndex);

                    String headerValue = headerCell.getStringCellValue();
                    String cellValue;
                    cellValue = obtenerValorVisibleCelda(currentCell);
                    rowMap.put(headerValue, cellValue);
                }
                countRows++;

                if (countRows % ROWS_PER_BATCH == 0) {
                    runtime();
                    Thread.sleep(200);
                }

                showProgressBarPerQuantity(countRows, totalRows);

                data.add(rowMap);
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return data;
    }

    public static class functions {
        public static Map<String, String> calcularSumaPorValoresUnicos(String filePath, String firstHeader, String secondHeader, int percent) throws IOException {
            List<Map<String, String>> data = leerExcel(filePath);
            Map<String, Double> sumaPorValorUnico = new HashMap<>();

            for (Map<String, String> row : data) {
                String firstHeaderValue = row.get(firstHeader);
                String secondHeaderValue = row.get(secondHeader);

                if (firstHeaderValue != null && secondHeaderValue != null) {
                    try {
                        double secondValue = Double.parseDouble(secondHeaderValue);
                        double porcentaje = (double) percent / 100;
                        double secondValueP = secondValue * porcentaje;

                        if (sumaPorValorUnico.containsKey(firstHeaderValue)) {
                            sumaPorValorUnico.put(firstHeaderValue, sumaPorValorUnico.get(firstHeaderValue) + (secondValueP));
                        } else {
                            sumaPorValorUnico.put(firstHeaderValue, secondValueP);
                        }
                    } catch (NumberFormatException e) {
                        // Ignora las filas que no tienen valores numéricos en el segundo encabezado
                    }
                }
            }

            // Redondea los valores a dos decimales
            Map<String, String> resultadoFormateado = new HashMap<>();
            DecimalFormat df = new DecimalFormat("#,##0.00");
            for (Map.Entry<String, Double> entry : sumaPorValorUnico.entrySet()) {
                double valor = entry.getValue();
                String valorFormateado = df.format(valor);
                resultadoFormateado.put(entry.getKey(), valorFormateado);
            }

            if (resultadoFormateado != null){
                return resultadoFormateado;
            } else {
                return null;
            }


        }

        public static Map<String, String> calcularSumaPorValoresUnicos(String filePath, String firstHeader, String secondHeader) throws IOException, InterruptedException {
            List<Map<String, String>> data = leerExcel(filePath);
            Map<String, Double> sumaPorValorUnico = new HashMap<>();

            int count = 0;
            int count2 = 0;
            //int rowsPerBatch = 5000;
            System.out.println("\n CALCULANDO SUMATORIA DE VALORES");
            for (Map<String, String> row : data) {
                String firstHeaderValue = row.get(firstHeader);
                String secondHeaderValue = row.get(secondHeader);

                if (firstHeaderValue != null && secondHeaderValue != null) {
                    try {
                        double secondValue = Double.parseDouble(secondHeaderValue);

                        if (sumaPorValorUnico.containsKey(firstHeaderValue)) {
                            sumaPorValorUnico.put(firstHeaderValue, sumaPorValorUnico.get(firstHeaderValue) + (secondValue));
                        } else {
                            sumaPorValorUnico.put(firstHeaderValue, secondValue);
                        }


                    } catch (NumberFormatException e) {
                        // Ignora las filas que no tienen valores numéricos en el segundo encabezado
                    }
                    count++;
                }
            }
            if (count % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            System.out.println("\n TERMINANDO PROCESO DE SUMATORIA DE VALORES");
            // Redondea los valores a dos decimales
            Map<String, String> resultadoFormateado = new HashMap<>();
            DecimalFormat df = new DecimalFormat("#,##0.00");
            int filasFinales = sumaPorValorUnico.size();
            for (Map.Entry<String, Double> entry : sumaPorValorUnico.entrySet()) {
                double valor = entry.getValue();
                String valorFormateado = df.format(valor);
                resultadoFormateado.put(entry.getKey(), valorFormateado);
                count2++;
            }
            if (count2 % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            showProgressBarPerQuantity(count2, filasFinales);


            if (resultadoFormateado != null) {
                return resultadoFormateado;
            }else {
                return null;
            }
        }

        public static Map<String, String> calcularConteoPorValoresUnicos(String filePath, String firstHeader, String secondHeader) throws IOException, InterruptedException {
            List<Map<String, String>> data = leerExcel(filePath);
            Map<String, Integer> conteoPorValorUnico = new HashMap<>();

            int count = 0;
            int count2 = 0;
            //int rowsPerBatch = 5000;

            for (Map<String, String> row : data) {
                String firstHeaderValue = row.get(firstHeader);

                if (firstHeaderValue != null) {
                    if (conteoPorValorUnico.containsKey(firstHeaderValue)) {
                        conteoPorValorUnico.put(firstHeaderValue, conteoPorValorUnico.get(firstHeaderValue) + 1);
                    } else {
                        conteoPorValorUnico.put(firstHeaderValue, 1);
                    }

                    count++;
                }
            }
            if (count % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            // Convertir el conteo a String para devolver un Map<String, String>
            Map<String, String> resultadoFormateado = new HashMap<>();
            for (Map.Entry<String, Integer> entry : conteoPorValorUnico.entrySet()) {
                resultadoFormateado.put(entry.getKey(), String.valueOf(entry.getValue()));
                count2++;
            }

            if (count2 % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            if (resultadoFormateado != null) {
                return resultadoFormateado;
            }else {
                return null;
            }
        }

        public static Map<String, String> calcularPromedioPorValoresUnicos(String filePath, String firstHeader, String secondHeader) throws IOException, InterruptedException {
            List<Map<String, String>> data = leerExcel(filePath);
            Map<String, Double> sumaPorValorUnico = new HashMap<>();
            Map<String, Integer> contadorPorValorUnico = new HashMap<>();

            int count = 0;
            int count2 = 0;
            //int rowsPerBatch = 5000;

            for (Map<String, String> row : data) {
                String firstHeaderValue = row.get(firstHeader);
                String secondHeaderValue = row.get(secondHeader);

                if (firstHeaderValue != null && secondHeaderValue != null) {
                    try {
                        double secondValue = Double.parseDouble(secondHeaderValue);

                        if (sumaPorValorUnico.containsKey(firstHeaderValue)) {
                            sumaPorValorUnico.put(firstHeaderValue, sumaPorValorUnico.get(firstHeaderValue) + secondValue);
                            contadorPorValorUnico.put(firstHeaderValue, contadorPorValorUnico.get(firstHeaderValue) + 1);
                        } else {
                            sumaPorValorUnico.put(firstHeaderValue, secondValue);
                            contadorPorValorUnico.put(firstHeaderValue, 1);
                        }
                    } catch (NumberFormatException e) {
                        // Ignora las filas que no tienen valores numéricos en el segundo encabezado
                    }
                    count++;
                }
            }
            if (count % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            Map<String, String> promedioPorValorUnico = new HashMap<>();
            DecimalFormat df = new DecimalFormat("#,##0.00");

            for (Map.Entry<String, Double> entry : sumaPorValorUnico.entrySet()) {
                String key = entry.getKey();
                double suma = entry.getValue();
                int contador = contadorPorValorUnico.get(key);
                double promedio = suma / contador;

                promedioPorValorUnico.put(key, df.format(promedio));
                count2++;
            }
            if (count2 % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            if (promedioPorValorUnico != null) {
                return promedioPorValorUnico;
            } else {
                return null;
            }
        }

        public static Map<String, String> calcularMinimoPorValoresUnicos(String filePath, String firstHeader, String secondHeader) throws IOException, InterruptedException {
            List<Map<String, String>> data = leerExcel(filePath);
            Map<String, Double> minimoPorValorUnico = new HashMap<>();

            int count = 0;
            int count2 = 0;
            //int rowsPerBatch = 5000;

            for (Map<String, String> row : data) {
                String firstHeaderValue = row.get(firstHeader);
                String secondHeaderValue = row.get(secondHeader);

                if (firstHeaderValue != null && secondHeaderValue != null) {
                    try {
                        double secondValue = Double.parseDouble(secondHeaderValue);

                        if (minimoPorValorUnico.containsKey(firstHeaderValue)) {
                            double currentMin = minimoPorValorUnico.get(firstHeaderValue);
                            if (secondValue < currentMin) {
                                minimoPorValorUnico.put(firstHeaderValue, secondValue);
                            }
                        } else {
                            minimoPorValorUnico.put(firstHeaderValue, secondValue);
                        }
                    } catch (NumberFormatException e) {
                        // Ignora las filas que no tienen valores numéricos en el segundo encabezado
                    }
                    count++;
                }
            }
            if (count % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            Map<String, String> minimoFormateado = new HashMap<>();
            DecimalFormat df = new DecimalFormat("#,##0.00");

            for (Map.Entry<String, Double> entry : minimoPorValorUnico.entrySet()) {
                minimoFormateado.put(entry.getKey(), df.format(entry.getValue()));
                count2++;
            }
            if (count2 % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            if (minimoFormateado != null) {
                return minimoFormateado;
            }else {
                return null;
            }
        }

        public static Map<String, String> calcularMaximoPorValoresUnicos(String filePath, String firstHeader, String secondHeader) throws IOException, InterruptedException {
            List<Map<String, String>> data = leerExcel(filePath);
            Map<String, Double> maximoPorValorUnico = new HashMap<>();

            int count = 0;
            int count2 = 0;
            //int rowsPerBatch = 5000;

            for (Map<String, String> row : data) {
                String firstHeaderValue = row.get(firstHeader);
                String secondHeaderValue = row.get(secondHeader);

                if (firstHeaderValue != null && secondHeaderValue != null) {
                    try {
                        double secondValue = Double.parseDouble(secondHeaderValue);

                        if (maximoPorValorUnico.containsKey(firstHeaderValue)) {
                            double currentMax = maximoPorValorUnico.get(firstHeaderValue);
                            if (secondValue > currentMax) {
                                maximoPorValorUnico.put(firstHeaderValue, secondValue);
                            }
                        } else {
                            maximoPorValorUnico.put(firstHeaderValue, secondValue);
                        }
                    } catch (NumberFormatException e) {
                        // Ignora las filas que no tienen valores numéricos en el segundo encabezado
                    }
                    count++;
                }
            }
            if (count % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            Map<String, String> maximoFormateado = new HashMap<>();
            DecimalFormat df = new DecimalFormat("#,##0.00");

            for (Map.Entry<String, Double> entry : maximoPorValorUnico.entrySet()) {
                maximoFormateado.put(entry.getKey(), df.format(entry.getValue()));
                count2++;
            }
            if (count2 % ROWS_PER_BATCH == 0) {
                runtime();
                Thread.sleep(200);
            }

            if (maximoFormateado != null) {
                return maximoFormateado;
            }else {
                return null;
            }
        }

    }



    /*------------------------------------------------------------------------------------------------------------------------------*/
    /*LECTURA DEL ARCHIVO MAESTRO PARA ANÁLISIS*/

    public static List<String> getHeaders(String excelFilePath, String sheetName) {
        List<String> headers = new ArrayList<>();
        try {
            Workbook workbook = createWorkbook(excelFilePath);
            Sheet sheet = workbook.getSheet(sheetName);
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                headers.add(obtenerValorVisibleCelda(cell));
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return headers;
    }


    /*public static List<Map<String, String>> obtenerValoresEncabezados2(String azureFile, String masterFile, String hoja, String fechaCorte) {
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
            Workbook workbook = createWorkbook(azureFile);
            Workbook workbook2 = createWorkbook(masterFile);
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
            String sht2 = "";
            List<String> dataList = createDualDropDownListsAndReturnSelectedValues(nameSheets1, nameSheets2);
            List<String> encabezados1 = null;
            List<String> encabezados2 = null;
            String encabezado = "";

            for (String seleccion : dataList) {
                String[] elementos = seleccion.split(SPECIAL_CHAR);

                sht1 = elementos[0];
                sht2 = elementos[1];
                System.out.println("ELEMENTOS SELECCIONADOS: " + sht1 + ", " + sht2);

                sheet2 = workbook2.getSheet(sht2);
            }
            sheet1 = workbook.getSheet(sht1);
            encabezados1 = getHeadersN(sheet1);

            JOptionPane.showMessageDialog(null, "Del siguiente menú escoja el primer encabezado ubicado en las hojas del archivo Maestro");
            assert encabezados1 != null;
            encabezado = mostrarMenu(encabezados1);

            encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezados1);

            JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" que será analizado");
            String codigo = mostrarMenu(encabezados2);
            JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda a la \"Fecha de corte\" que será analizada");
            String fechaCorteMF = mostrarMenu(encabezados2);


            for (String seleccion : dataList) {
                if (sht2.equals(hoja)) {
                    if (!fechaCorte.equals(fechaCorteMF)) {
                        errorMessage("Por favor verifique que los encabezados correspondientes a las fechas" +
                                "\n tengan un formato tipo FECHA idéntica a " + fechaCorte);

                        errorMessage("No es posible completar el análisis de la hoja [" + hoja +
                                "]\n el formato de fecha no es el correcto");
                    } else {
                        valoresEncabezados2 = obtenerValoresPorFilas(workbook, workbook2, sht1, sht2, codigo, fechaCorteMF, headers);
                        mapList = createMapList(valoresEncabezados2, codigo, fechaCorteMF);
                        for (Map<String, String> map : mapList) {
                            System.out.println("Analizando valores... ");
                            for (Map.Entry<String, String> entry : map.entrySet()) {
                                System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
                            }
                        }
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
    }*/


    /*-----------------------------------------------------------------------------------------------------------------------------------------*/
    public static List<String> getHeadersN(Sheet sheet) {
        List<String> columnNames = new ArrayList<>();
        Row row = sheet.getRow(0);
        try {
            System.out.println("PROCESANDO CAMPOS...");
            for (Iterator<Cell> it = row.cellIterator(); it.hasNext(); ) {
                Cell cell = it.next();
                columnNames.add(obtenerValorVisibleCelda(cell));
                //System.out.println(obtenerValorVisibleCelda(cell));
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return columnNames;
    }

    public static List<String> getHeadersN(String filePath, String sheetName) {
        List<String> columnNames = new ArrayList<>();

        try /*(Workbook workbook = WorkbookFactory.create(new File(filePath)))*/ {
            System.out.println("AQUI ESTOY");
            Workbook workbook = createWorkbook(filePath);
            Sheet sheet = workbook.getSheet(sheetName);
            Row row = sheet.getRow(0);
            System.out.println("PROCESANDO CAMPOS...");
            for (Iterator<Cell> it = row.cellIterator(); it.hasNext(); ) {
                Cell cell = it.next();
                columnNames.add(obtenerValorVisibleCelda(cell));
                //System.out.println(obtenerValorVisibleCelda(cell));
            }
            workbook.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return columnNames;
    }

    /*public static Map<String, Object> getHeaderValuesN(Sheet sheet, List<String> headers){
        Map<String, Object> rowData = new HashMap<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        try {
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }
                Iterator<String> columnNameIterator = headers.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String columnName = columnNameIterator.next();
                    String value = "";
                    if (cell != null) {
                        //System.out.println("The cell contains a numeric value." + cell.getCellType());
                        value = obtenerValorVisibleCelda(cell);
                        System.out.println("VALUE: " + value + ", " + cell.getRowIndex());
                        rowData.put(columnName, value);
                    }
                    runtime();
                    Thread.sleep(200);
                    //waitSeconds(2);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }


        return rowData;
    }*/

    public static List<Map<String, Object>> getHeaderValuesN(Sheet sheet, List<String> headers) {
        List<Map<String, Object>> dataList = new ArrayList<>();

        Row row;
        Iterator<Row> rowIterator = sheet.iterator();

        try {

            int currentRow = 0;
            //int rowsPerBatch = 5000;
            int totalRows = sheet.getPhysicalNumberOfRows() - 1;

            System.out.println("PROCESANDO VALORES");

            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                Iterator<String> columnNameIterator = headers.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();
                Map<String, Object> rowData = new HashMap<>();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String columnName = columnNameIterator.next();
                    Object value;

                    if (cell != null) {
                        // Obtener el valor de la celda
                        value = obtenerValorVisibleCelda(cell);
                        rowData.put(columnName, value);
                    }
                    runtime();
                    Thread.sleep(200);
                }

                dataList.add(rowData);
                currentRow++;

                if (currentRow % ROWS_PER_BATCH == 0) {
                    runtime();
                    Thread.sleep(200);
                }

                showProgressBarPerQuantity(currentRow, totalRows);


            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return dataList;
    }


    public static List<Map<String, Object>> getHeaderFilterValuesNS(Sheet sheet, List<String> headers, String campoFiltrar, String valorIni, String valorFin) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }
                int campoFiltrarIndex = headers.indexOf(campoFiltrar);
                if (campoFiltrarIndex == -1) {
                    System.err.println("El campo especificado para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex));
                Iterator<String> columnNameIterator = headers.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();
                if (valueCampoFiltrar.compareTo(valorIni) >= 0 && valueCampoFiltrar.compareTo(valorFin) <= 0) {

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }

                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }

                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    public static List<Map<String, Object>> getHeaderFilterValuesNN(Sheet sheet, List<String> headers, String campoFiltrar, double valorIni, double valorFin) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 3000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }
                int campoFiltrarIndex = headers.indexOf(campoFiltrar);
                if (campoFiltrarIndex == -1) {
                    System.err.println("El campo especificado para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex));
                valueCampoFiltrar = valueCampoFiltrar.replaceAll("\\.", "");
                valueCampoFiltrar = valueCampoFiltrar.replaceAll(",", ".");
                double numericValueCampoFiltrar = parseDouble(valueCampoFiltrar);

                Iterator<String> columnNameIterator = headers.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();
                if (numericValueCampoFiltrar >= valorIni && numericValueCampoFiltrar <= valorFin) {

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }

                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    if (currentRow % ROWS_PER_BATCH == 0) {
                        //System.out.println("3000 registros");
                        runtime();
                        Thread.sleep(200);
                    }

                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    private static double parseDouble(String value) {
        try {
            return Double.parseDouble(value);
        } catch (NumberFormatException e) {
            // Handle the case where the value is not a valid double
            return 0.0;  // Change this to the default value you want
        }
    }


    public static List<Map<String, Object>> getHeaderFilterValuesNSS(Sheet sheet, List<String> headers, String campoFiltrar1, String valorIni1, String valorFin1, String campoFiltrar2, String valorIni2, String valorFin2) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
                int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
                if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                    System.err.println("Al menos uno de los campos especificados para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar1 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex1));
                String valueCampoFiltrar2 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex2));

                if ((valueCampoFiltrar1.compareTo(valorIni1) >= 0 && valueCampoFiltrar1.compareTo(valorFin1) <= 0) &&
                        (valueCampoFiltrar2.compareTo(valorIni2) >= 0 && valueCampoFiltrar2.compareTo(valorFin2) <= 0)) {

                    Iterator<String> columnNameIterator = headers.iterator();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }


                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    public static List<Map<String, Object>> getHeaderFilterValuesNSN(Sheet sheet, List<String> headers, String campoFiltrar1, String valorIni1, String valorFin1, String campoFiltrar2, double valorIni2, double valorFin2) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
                int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
                if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                    System.err.println("Al menos uno de los campos especificados para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar1 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex1));
                String valueCampoFiltrar2 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex2));
                valueCampoFiltrar2 = valueCampoFiltrar2.replaceAll("\\.", "");
                valueCampoFiltrar2 = valueCampoFiltrar2.replace(",", "");
                double numeroCampoFiltrar2 = Double.parseDouble(valueCampoFiltrar2);

                if ((valueCampoFiltrar1.compareTo(valorIni1) >= 0 && valueCampoFiltrar1.compareTo(valorFin1) <= 0) &&
                        (numeroCampoFiltrar2 >= valorIni2 && numeroCampoFiltrar2 <= valorFin2)) {

                    Iterator<String> columnNameIterator = headers.iterator();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;
                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }


                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    public static List<Map<String, Object>> getHeaderFilterValuesNSD(Sheet sheet, List<String> headers, String campoFiltrar1, String valorIni1, String valorFin1, String campoFiltrar2, Date valorIni2, Date valorFin2) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
                int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
                if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                    System.err.println("Al menos uno de los campos especificados para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar1 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex1));
                Date valueCampoFiltrar2 = parseDate(obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex2)));


                if ((valueCampoFiltrar1.compareTo(valorIni1) >= 0 && valueCampoFiltrar1.compareTo(valorFin1) <= 0) &&
                        (valueCampoFiltrar2.compareTo(valorIni2) >= 0 && valueCampoFiltrar2.compareTo(valorFin2) <= 0) &&
                        (valueCampoFiltrar1 != null && valueCampoFiltrar2 != null)) {

                    Iterator<String> columnNameIterator = headers.iterator();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;
                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }


                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }


    public static List<Map<String, Object>> getHeaderFilterValuesNNN(Sheet sheet, List<String> headers, String campoFiltrar1, double valorIni1, double valorFin1, String campoFiltrar2, double valorIni2, double valorFin2) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
                int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);

                if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                    System.err.println("El campo especificado para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar1 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex1));
                String valueCampoFiltrar2 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex2));
                valueCampoFiltrar1 = valueCampoFiltrar1.replaceAll("\\.", "");
                valueCampoFiltrar1 = valueCampoFiltrar1.replace(",", ".");
                valueCampoFiltrar2 = valueCampoFiltrar2.replaceAll("\\.", "");
                valueCampoFiltrar2 = valueCampoFiltrar2.replace(",", ".");


                double numericCampoFiltrar1 = Double.parseDouble(valueCampoFiltrar1);
                double numericCampoFiltrar2 = Double.parseDouble(valueCampoFiltrar2);

                if ((numericCampoFiltrar1 >= valorIni1 && numericCampoFiltrar1 <= valorFin1) &&
                        (numericCampoFiltrar2 >= valorIni2 && numericCampoFiltrar2 <= valorFin2)) {

                    Iterator<String> columnNameIterator = headers.iterator();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }


                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    public static List<Map<String, Object>> getHeaderFilterValuesNND(Sheet sheet, List<String> headers, String campoFiltrar1, double valorIni1, double valorFin1, String campoFiltrar2, Date valorIni2, Date valorFin2) {

        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);
        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
                int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);

                if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                    System.err.println("El campo especificado para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar1 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex1));
                valueCampoFiltrar1 = valueCampoFiltrar1.replaceAll("\\.", "");
                valueCampoFiltrar1 = valueCampoFiltrar1.replace(",", "");

                double numericCampoFiltrar1 = Double.parseDouble(valueCampoFiltrar1);
                Date valueCampoFiltrar2 = parseDate(obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex2)));

                if ((numericCampoFiltrar1 >= valorIni1 && numericCampoFiltrar1 <= valorFin1) &&
                        (valueCampoFiltrar2.after(valorIni2) || valueCampoFiltrar2.equals(valorIni2)) &&
                        (valueCampoFiltrar2.before(valorFin2) || valueCampoFiltrar2.equals(valorFin2))) {

                    Iterator<String> columnNameIterator = headers.iterator();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }

                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    private static Date parseDate(String dateStr) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
        try {
                return dateFormat.parse(dateStr);

        } catch (ParseException e) {
            System.err.println("Información incompleta: " + e.getMessage());
            return null;
        }
    }

    public static Date parsearFecha(String fechaStr) {
        String[] patrones = {"dd/MM/yyyy", "yyyyMMdd", "MMddyyyy", "yyyyMMddHHmmss"};

        for (String patron : patrones) {
            SimpleDateFormat formato = new SimpleDateFormat(patron);
            try {
                if (fechaStr != null) {
                    return formato.parse(fechaStr);
                }
            } catch (ParseException e) {
                // Ignorar la excepción y probar con el siguiente patrón
            }
        }

        System.err.println("Error parsing date: Unparseable date: " + fechaStr);
        return null;
    }

    public static List<Map<String, Object>> getHeaderFilterValuesNNS(Sheet sheet, List<String> headers, String campoFiltrar1, double valorIni1, double valorFin1, String campoFiltrar2, String valorIni2, String valorFin2) {
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            //int rowsPerBatch = 5000;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }

                int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
                int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);

                if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                    System.err.println("El campo especificado para el filtro no existe");
                    return datosFiltrados;
                }

                String valueCampoFiltrar1 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex1));
                valueCampoFiltrar1 = valueCampoFiltrar1.replaceAll("\\.", "");
                valueCampoFiltrar1 = valueCampoFiltrar1.replace(",", "");
                String valueCampoFiltrar2 = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex2));
                valueCampoFiltrar2 = valueCampoFiltrar2.replaceAll("\\.", "");
                valueCampoFiltrar2 = valueCampoFiltrar2.replace(",", "");

                double numeroCampoFiltrar1 = Double.parseDouble(valueCampoFiltrar1);
                double numeroCampoFiltrar2 = Double.parseDouble(valueCampoFiltrar2);

                if ((numeroCampoFiltrar1 >= valorIni1 && numeroCampoFiltrar1 <= valorFin1) &&
                        (numeroCampoFiltrar2 >= Double.parseDouble(valorIni2) && numeroCampoFiltrar2 <= Double.parseDouble(valorFin2))) {

                    Iterator<String> columnNameIterator = headers.iterator();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            value = obtenerValorVisibleCelda(cell);
                            rowData.put(columnName, value);
                        }
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    if (currentRow % ROWS_PER_BATCH == 0) {
                        runtime();
                        Thread.sleep(200);
                    }

                    showProgressBarPerQuantity(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    public static void showProgressBarPerQuantity(int current, int total) {
        int progressBarWidth = 50;
        int progress = (int) ((double) current / total * progressBarWidth);

        StringBuilder progressBar = new StringBuilder("[");
        for (int i = 0; i < progressBarWidth; i++) {
            if (i < progress) {
                progressBar.append("=");
            } else {
                progressBar.append(" ");
            }
        }
        progressBar.append("] " + current + "/" + total);
        System.out.print("\r" + progressBar.toString());
    }

    public static void logWinsToFile(String filePath, List<String> messages) {
        writeExcelFile(filePath, messages, "messages");
    }

    public static void logErrorsToFile(String filePath, List<String> errors) {
        writeExcelFile(filePath, errors, "errors");
    }

    public static void writeExcelFile(String filePath, List<String> messages, String folderName) {
        // Obtener el nombre del archivo sin la extensión
        String fileName = new File(filePath).getName();
        String folderPath = filePath.replace(fileName, folderName);
        System.err.println(folderPath);

        // Crear la carpeta si no existe
        File folder = new File(folderPath);
        folder.mkdirs();

        System.err.println(folderName);
        System.err.println(fileName);

        String formattedDate = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss").format(new Date());
        String excelFileName = fileName.replace(".xlsx", "-" + folderName + "-" + formattedDate + ".xlsx");

        System.err.println(excelFileName);

        // Agregar "-estatus" al nombre del archivo Excel
        String excelFilePath = folderPath + File.separator + excelFileName;

        try (Workbook workbook = new XSSFWorkbook(); // Utiliza XSSFWorkbook para archivos .xlsx
             FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {

            Sheet sheet = workbook.createSheet("LogSheet");

            int rowNum = 0;
            for (String message : messages) {
                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(0);
                cell.setCellValue(message);
            }

            workbook.write(fileOut);
            System.out.println("Mensajes registrados en el archivo Excel: " + excelFilePath);
        } catch (IOException e) {
            // Manejar cualquier excepción de IO, por ejemplo, imprimir en la consola
            e.printStackTrace();
        }
    }


    /*-----------------------------------------------------------------------------------------------------------------------------------------*/
    public static final String SPECIAL_CHAR = " -X- ";


    public static List<String> createDualDropDownListsAndReturnSelectedValues(List<String> list1, List<String> list2) {
        List<String> selectedValues = new ArrayList<>();

        JFrame frame = new JFrame("SELECCIÓN DE HOJAS");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 300);
        frame.setLayout(new FlowLayout());

        JComboBox<String> dropdown1 = new JComboBox<>(list1.toArray(new String[0]));
        JComboBox<String> dropdown2 = new JComboBox<>(list2.toArray(new String[0]));
        JButton addButton = new JButton("Agregar Selecciones");

        frame.add(dropdown1);
        frame.add(dropdown2);
        frame.add(addButton);

        DefaultListModel<String> listModel = new DefaultListModel<>();
        JList<String> selectionsList = new JList<>(listModel);
        JScrollPane scrollPane = new JScrollPane(selectionsList);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        frame.add(scrollPane);

        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String selectedValue1 = (String) dropdown1.getSelectedItem();
                String selectedValue2 = (String) dropdown2.getSelectedItem();

                if (selectedValue1 != null && selectedValue2 != null) {
                    String combinedSelection = selectedValue1 + SPECIAL_CHAR + selectedValue2;
                    selectedValues.add(combinedSelection);

                    listModel.addElement(combinedSelection);

                    // Eliminar elementos seleccionados de los desplegables
                    list1.remove(selectedValue1);
                    list2.remove(selectedValue2);

                    // Actualizar los modelos de los desplegables
                    dropdown1.setModel(new DefaultComboBoxModel<>(list1.toArray(new String[0])));
                    dropdown2.setModel(new DefaultComboBoxModel<>(list2.toArray(new String[0])));

                    System.out.println("Elementos agregados: " + combinedSelection);
                } else {
                    // Puedes mostrar un mensaje de error si ambos elementos no están seleccionados
                    JOptionPane.showMessageDialog(frame, "Selecciona un elemento de cada lista", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Botón para eliminar selecciones marcadas
        JButton removeButton = new JButton("Eliminar Selecciones");
        frame.add(removeButton);

        removeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Eliminar selecciones marcadas
                int[] selectedIndices = selectionsList.getSelectedIndices();
                for (int i = selectedIndices.length - 1; i >= 0; i--) {
                    String removedValue = listModel.getElementAt(selectedIndices[i]);
                    selectedValues.remove(removedValue);

                    // Recuperar elementos eliminados a los desplegables
                    String[] parts = removedValue.split(SPECIAL_CHAR);
                    if (!list1.contains(parts[0])) {
                        list1.add(parts[0]);
                    }
                    if (!list2.contains(parts[1])) {
                        list2.add(parts[1]);
                    }

                    listModel.removeElementAt(selectedIndices[i]);
                }

                // Actualizar los modelos de los desplegables
                dropdown1.setModel(new DefaultComboBoxModel<>(list1.toArray(new String[0])));
                dropdown2.setModel(new DefaultComboBoxModel<>(list2.toArray(new String[0])));
            }
        });

        // Botón para terminar el proceso de selección
        JButton finishButton = new JButton("Terminar Selección");
        frame.add(finishButton);

        finishButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Puedes realizar acciones finales aquí, por ejemplo, cerrar la aplicación
                frame.dispose();
            }
        });

        frame.setVisible(true);

        // Esperar hasta que se cierre la ventana
        while (frame.isVisible()) {
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return selectedValues;
    }



    public static void errorMessage(String mensaje) {
        JLabel label = new JLabel("<html><font color='red'>" + mensaje + "</font></html>");
        label.setFont(new Font("Arial", Font.PLAIN, 14)); // Puedes ajustar la fuente según tus preferencias

        JOptionPane.showMessageDialog(null, label, "Error", JOptionPane.ERROR_MESSAGE);
    }


    public static String mostrarMenu(List<String> opciones) {
        List<String> opcionesConNinguno = new ArrayList<>(opciones);
        opcionesConNinguno.add(0, "Ninguno");

        JFrame frame = new JFrame("Menú de Opciones");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JComboBox<String> comboBox = new JComboBox<>(opcionesConNinguno.toArray(new String[0]));
        comboBox.setSelectedIndex(0);

        JButton button = new JButton("Seleccionar");

        ActionListener actionListener = e -> frame.dispose();

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


    public static String mostrarCuadroDeTexto() {
        // Crea una ventana Swing
        JFrame frame = new JFrame("Ingrese el valor Indicado");

        // Crea un cuadro de texto
        JTextField textField = new JTextField(20); // 20 es el ancho del cuadro de texto

        // Crea un botón
        JButton button = new JButton("Ingresar");

        // Crea una variable para almacenar el texto ingresado
        AtomicReference<String> textoIngresado = new AtomicReference<>("");

        // Crea un objeto de tipo Semaphore para bloquear hasta que se ingrese el texto
        Semaphore semaphore = new Semaphore(0);

        // Agrega un ActionListener al botón para manejar el evento de clic
        button.addActionListener(e -> {
            textoIngresado.set(textField.getText());
            semaphore.release(); // Libera el semáforo para indicar que se ingresó el texto
            frame.dispose();
        });

        // Crea un panel y agrega el cuadro de texto y el botón a él
        JPanel panel = new JPanel();
        panel.add(textField);
        panel.add(button);

        // Agrega el panel a la ventana
        frame.add(panel);

        // Configura las propiedades de la ventana
        frame.setSize(300, 100); // Tamaño de la ventana
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true); // Hace visible la ventana

        try {
            semaphore.acquire(); // Bloquea hasta que se libere el semáforo (se ingrese el texto)
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        return textoIngresado.get();
    }

    private static final int FRAME_WIDTH = 320;
    private static final int FRAME_HEIGHT = 100;
    public static String showDateChooser() {
        JFrame frame = new JFrame("Seleccionar Fecha");
        frame.setPreferredSize(new Dimension(FRAME_WIDTH, FRAME_HEIGHT));
        JDateChooser dateChooser = new JDateChooser();
        dateChooser.setFont(new Font("Arial", Font.PLAIN, 18));
        JButton okButton = new JButton("Aceptar");

        okButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                frame.dispose(); // Cerrar la ventana después de hacer clic en "Aceptar"
            }
        });

        frame.setLayout(new BoxLayout(frame.getContentPane(), BoxLayout.Y_AXIS));
        frame.add(dateChooser);
        frame.add(okButton);

        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);

        // Esperar hasta que se cierre la ventana
        while (frame.isVisible()) {
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        return sdf.format(dateChooser.getDate());
    }

    public static String showMonthYearChooser() {
        JFrame frame = new JFrame("Seleccionar Mes y Año");
        frame.setPreferredSize(new Dimension(FRAME_WIDTH, FRAME_HEIGHT));
        JDateChooser dateChooser = new JDateChooser();
        dateChooser.setFont(new Font("Arial", Font.PLAIN, 18));
        dateChooser.setDateFormatString("MM/yyyy"); // Establecer el formato para mostrar solo mes y año
        JButton okButton = new JButton("Aceptar");

        okButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                frame.dispose(); // Cerrar la ventana después de hacer clic en "Aceptar"
            }
        });

        frame.setLayout(new BoxLayout(frame.getContentPane(), BoxLayout.Y_AXIS));
        frame.add(dateChooser);
        frame.add(okButton);

        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);

        // Esperar hasta que se cierre la ventana
        while (frame.isVisible()) {
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        SimpleDateFormat sdf = new SimpleDateFormat("MM/yyyy");
        return sdf.format(dateChooser.getDate());
    }

    /*public static void showProgressBarPercent(int current, int total) {
        int progressBarWidth = 50;
        int progress = (int) ((double) current / total * 100);

        StringBuilder progressBar = new StringBuilder("[");
        for (int i = 0; i < progressBarWidth; i++) {
            if (i < progress * progressBarWidth / 100) {
                progressBar.append("||");
            } else {
                progressBar.append(" ");
            }
        }
        progressBar.append("] " + progress + "%");
        System.out.print("\r" + progressBar.toString());
    }*/

}




