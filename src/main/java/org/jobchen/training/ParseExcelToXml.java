package org.jobchen.training;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.jamesmurty.utils.XMLBuilder2;
import edu.npu.fastexcel.ExcelException;
import edu.npu.fastexcel.FastExcel;
import lombok.Data;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.w3c.dom.Document;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

/**
 * @author : jobchen
 * @version : 1.0.0 Copyright (c) 2021 jobchen, All Rights Reserved.
 * @program : Training
 * @description : <br>Parse excel data to xml</br>
 * Use xmlbuilder as xml parser reference: https://github.com/jmurty/java-xmlbuilder
 * @create : 2021-01-12 55:29
 * @history : Who--->When--->What
 */
public class ParseExcelToXml {

    public static void main(String[] args) throws IOException, TransformerException, ExcelException {

        System.out.println("Start generate xml");

        long startTime = System.currentTimeMillis();

        parse("/Users/jobchen/Desktop/sample001.xlsx", 1, 4, 2, 3, "/Users/jobchen/Desktop");
        System.out.println("Finish generate xml, cost: " + (System.currentTimeMillis() - startTime) + "(ms)");

//        XMLBuilder2 builder = XMLBuilder2.create("Projects");
//        builder.xpathFind("//Projects").e("java-xmlbuilder");
//        builder.xpathFind("//java-xmlbuilder").e("Location");
//        builder.xpathFind("//Location").e("JetS3t");
//
//        xmlBuilderToFile(builder, "/Users/jobchen/Desktop", "xml_test2");
    }

    /**
     * Generate xml file according to excel
     *
     * @param excelPath            Excel file path
     * @param xmlNodePathColumnIdx Xml node path index, start from {@code 1}
     * @param dataColumnIdx        Data index, start from {@code 1}
     * @param xmlFileRelativePath  Generated xml file relative path, such as '/users/xml', xml file name is decided by xml root
     * @throws IOException
     * @throws TransformerException
     */
    private static void parse(String excelPath,
                              int xmlNodePathColumnIdx,
                              int dataColumnIdx,
                              int attributeNameColumnIdx,
                              int attributeValueColumnIdx,
                              String xmlFileRelativePath) throws IOException, TransformerException, ExcelException {

        // Collect root node name and a init root xml builder
        Map<String, XMLBuilder2> rootXmlBuilderMap = Maps.newHashMap();

        // Collect each row data, key is node path
//        Map<String, String> dataMap = initXmlBuilderAndCollectData(new File(excelPath), xmlNodePathColumnIdx, dataColumnIdx, rootXmlBuilderMap);
        // This one is fast, so recommend this.
        Map<String, XmlData> dataMap = initXmlBuilderAndCollectDataByFastExcel(
                new File(excelPath), attributeNameColumnIdx, attributeValueColumnIdx, xmlNodePathColumnIdx, dataColumnIdx, rootXmlBuilderMap);
        // Too old to use
//        Map<String, String> dataMap = initXmlBuilderAndCollectDataByFastExcel2(new File(excelPath), xmlNodePathColumnIdx, dataColumnIdx, rootXmlBuilderMap);

//        Map<String, String> dataMap = initXmlBuilderAndCollectDataByEasyExcel(new File(excelPath), xmlNodePathColumnIdx, dataColumnIdx, rootXmlBuilderMap);

        if (MapUtils.isNotEmpty(dataMap)) {

            fillRootXmlBuilder(rootXmlBuilderMap, dataMap);

            for (Map.Entry<String, XMLBuilder2> xmlBuilder2Entry : rootXmlBuilderMap.entrySet()) {
                xmlBuilderToFile(xmlBuilder2Entry.getValue(), xmlFileRelativePath, xmlBuilder2Entry.getKey());
            }
        }
    }

    /**
     * Fill xml builder leaf nodes and data
     *
     * @param rootXmlBuilderMap Initiated root xml builder which only has one root node
     * @param dataMap           Xml node data map
     */
    private static void fillRootXmlBuilder(Map<String, XMLBuilder2> rootXmlBuilderMap,
                                           Map<String, XmlData> dataMap) {
        String[] dataKeys;

        for (Map.Entry<String, XmlData> dataEntry : dataMap.entrySet()) {

            dataKeys = StringUtils.split(dataEntry.getKey(), "/");

            // Root node
            if (dataKeys.length == 1) {
                continue;
            }

            // Leaf node
            if (dataKeys.length > 1) {
                for (int i = 1; i < dataKeys.length; i++) {
                    if (i != dataKeys.length - 1) {
                        if (!checkIfExist(rootXmlBuilderMap.get(dataKeys[0]), dataKeys[i])) {
                            rootXmlBuilderMap.get(dataKeys[0])
                                    .xpathFind("//" + dataKeys[i - 1])
                                    .e(dataKeys[i])
                                    .a(dataEntry.getValue().getAttributeName(), dataEntry.getValue().getAttributeValue());
                        }
                    } else {
                        rootXmlBuilderMap.get(dataKeys[0])
                                .xpathFind("//" + dataKeys[i - 1])
                                .e(dataKeys[i])
                                .a(dataEntry.getValue().getAttributeName(), dataEntry.getValue().getAttributeValue())
                                .t(dataEntry.getValue().getData());
                    }
                }
            }

            System.out.println(dataEntry.getKey() + ": " + rootXmlBuilderMap.get(dataKeys[0]).asString());
        }
    }

    /**
     * Build xml builder precondition data
     *
     * @param excelFile            Excel file
     * @param xmlNodePathColumnIdx Xml path column index
     * @param dataColumnIdx        Xml node data column
     * @param rootXmlBuilderMap    Xml root node map
     * @throws IOException
     */
    private static Map<String, String> initXmlBuilderAndCollectData(File excelFile,
                                                                    int xmlNodePathColumnIdx,
                                                                    int dataColumnIdx,
                                                                    Map<String, XMLBuilder2> rootXmlBuilderMap) throws IOException {

        System.out.println("Start init xml builder data");

        long startTime = System.currentTimeMillis();

        ReadableWorkbook wb = new ReadableWorkbook(new FileInputStream(excelFile));

        org.dhatim.fastexcel.reader.Sheet firstSheet = wb.getFirstSheet();

        Workbook workbook = WorkbookFactory.create(excelFile);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " sheets : ");

        // Handle first sheet data
        Sheet sheet = workbook.getSheetAt(0);

        // Collect each row data, key is node path
        Map<String, String> dataMap = Maps.newHashMap();

        System.out.println("First sheet has " + sheet.getPhysicalNumberOfRows() + " rows: ");

        for (Row row : sheet) {

            // Cell value
            String cellValue;

            String dataKey = null;
            String data = null;

            // Record column of a row
            int cellIndex = 0;

            for (Cell cell : row) {

                cellValue = getCellValue(cell);
                cellValue = StringUtils.isBlank(cellValue) ? "" : cellValue.trim();

                // Collect root names from first column
                if (cellIndex == (xmlNodePathColumnIdx - 1)) {
                    cellValue = cellValue.replaceAll("/+", "/");
                    dataKey = cellValue;
                    String[] slashArrays = StringUtils.split(cellValue, "/");
                    if (slashArrays.length > 0 && !rootXmlBuilderMap.containsKey(slashArrays[0])) {
                        rootXmlBuilderMap.put(slashArrays[0], XMLBuilder2.create(slashArrays[0]));
                    }
                }

                if (cellIndex == (dataColumnIdx - 1)) {
                    data = cellValue;
                }

                cellIndex += 1;
            }

            if (null != dataKey) {
                if (dataMap.containsKey(dataKey)) {
                    System.out.println("Already exist xml path: " + dataKey + ", data: " + data + ", will replace with data: " + data);
                }
                dataMap.put(dataKey, data);
            }
        }

        // Closing the workbook
        workbook.close();

        System.out.println("Finish init xml builder data, cost: " + (System.currentTimeMillis() - startTime) + "(ms)");

        return dataMap;
    }

    /**
     * Build xml builder precondition data
     *
     * @param excelFile            Excel file
     * @param xmlNodePathColumnIdx Xml path column index
     * @param dataColumnIdx        Xml node data column
     * @param rootXmlBuilderMap    Xml root node map
     * @throws IOException
     */
    private static Map<String, XmlData> initXmlBuilderAndCollectDataByFastExcel(File excelFile,
                                                                                int attributeNameColumnIdx,
                                                                                int attributeValueColumnIdx,
                                                                                int xmlNodePathColumnIdx,
                                                                                int dataColumnIdx,
                                                                                Map<String, XMLBuilder2> rootXmlBuilderMap) throws IOException {

        System.out.println("Start init xml builder data by fast_excel");

        long startTime = System.currentTimeMillis();

        ReadableWorkbook workbook = new ReadableWorkbook(new FileInputStream(excelFile));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getSheets().count() + " sheets : ");

        // Handle first sheet
        org.dhatim.fastexcel.reader.Sheet firstSheet = workbook.getFirstSheet();

        List<org.dhatim.fastexcel.reader.Row> rows = firstSheet.read();

        // Collect each row data, key is node path
        Map<String, String> dataMap = Maps.newHashMap();

        Map<String, XmlData> xmlDataMap = Maps.newHashMap();

        System.out.println("First sheet has " + rows.size() + " rows: ");

        for (org.dhatim.fastexcel.reader.Row row : rows) {

            // Cell value
            String cellValue;

            String dataKey = null;
            String data = null;

            // Record column of a row
            int cellIndex = 0;

            XmlData xmlData = new XmlData();

            for (org.dhatim.fastexcel.reader.Cell cell : Lists.newArrayList(row.iterator())) {

                cellValue = cell.getRawValue();
                cellValue = StringUtils.isBlank(cellValue) ? "" : cellValue.trim();

                // Collect root names from first column
                if (cellIndex == (xmlNodePathColumnIdx - 1)) {

                    cellValue = cellValue.replaceAll("/+", "/");

                    xmlData.setPath(cellValue);

                    dataKey = cellValue;
                    String[] slashArrays = StringUtils.split(cellValue, "/");
                    if (slashArrays.length > 0 && !rootXmlBuilderMap.containsKey(slashArrays[0])) {
                        rootXmlBuilderMap.put(slashArrays[0], XMLBuilder2.create(slashArrays[0]));
                    }
                }

                if (cellIndex == (dataColumnIdx - 1)) {
                    data = cellValue;
                    xmlData.setData(data);
                }

                if (cellIndex == (attributeNameColumnIdx - 1)) {
                    xmlData.setAttributeName(cellValue);
                }

                if (cellIndex == (attributeValueColumnIdx - 1)) {
                    xmlData.setAttributeValue(cellValue);
                }

                cellIndex += 1;
            }

            if (null != dataKey) {
                if (dataMap.containsKey(dataKey)) {
                    System.out.println("Already exist xml path: " + dataKey + ", data: " + data + ", will replace with data: " + data);
                }
                dataMap.put(dataKey, data);
                xmlDataMap.put(dataKey, xmlData);
            }
        }

        System.out.println("Finish init xml builder data by fast_excel, cost: " + (System.currentTimeMillis() - startTime) + "(ms)");

        return xmlDataMap;
    }

    /**
     * <br>Build xml builder precondition data</br>
     * This is deprecated, cause can only parse xls of 93-2003
     *
     * @param excelFile            Excel file
     * @param xmlNodePathColumnIdx Xml path column index
     * @param dataColumnIdx        Xml node data column
     * @param rootXmlBuilderMap    Xml root node map
     * @throws IOException
     */
    @Deprecated
    private static Map<String, String> initXmlBuilderAndCollectDataByFastExcel2(File excelFile,
                                                                                int xmlNodePathColumnIdx,
                                                                                int dataColumnIdx,
                                                                                Map<String, XMLBuilder2> rootXmlBuilderMap) throws ExcelException {

        System.out.println("Start init xml builder data by fast_excel2");

        long startTime = System.currentTimeMillis();

        edu.npu.fastexcel.Workbook workbook = FastExcel.createReadableWorkbook(excelFile);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.sheetCount() + " sheets : ");

        // Handle first sheet
        edu.npu.fastexcel.Sheet firstSheet = workbook.getSheet(0);

        // Collect each row data, key is node path
        Map<String, String> dataMap = Maps.newHashMap();

        System.out.println("First sheet has " + (firstSheet.getLastRow() + 1) + " rows: ");

        for (int rowIndex = firstSheet.getFirstRow(); rowIndex < firstSheet.getLastRow(); rowIndex++) {

            // Cell value
            String cellValue;

            String dataKey = null;
            String data = null;

            for (int columnIndex = firstSheet.getFirstColumn(); columnIndex < firstSheet.getLastColumn(); columnIndex++) {
                cellValue = firstSheet.getCell(rowIndex, columnIndex);
                cellValue = StringUtils.isBlank(cellValue) ? "" : cellValue.trim();
                if (columnIndex == (xmlNodePathColumnIdx - 1)) {
                    cellValue = cellValue.replaceAll("/+", "/");
                    dataKey = cellValue;
                    String[] slashArrays = StringUtils.split(cellValue, "/");
                    if (slashArrays.length > 0 && !rootXmlBuilderMap.containsKey(slashArrays[0])) {
                        rootXmlBuilderMap.put(slashArrays[0], XMLBuilder2.create(slashArrays[0]));
                    }
                }

                if (columnIndex == (dataColumnIdx - 1)) {
                    data = cellValue;
                }
            }

            if (null != dataKey) {
                if (dataMap.containsKey(dataKey)) {
                    System.out.println("Already exist xml path: " + dataKey + ", data: " + data + ", will replace with data: " + data);
                }
                dataMap.put(dataKey, data);
            }
        }

        System.out.println("Finish init xml builder data by fast_excel2, cost: " + (System.currentTimeMillis() - startTime) + "(ms)");

        return dataMap;
    }

    /**
     * Build xml builder precondition data
     *
     * @param excelFile            Excel file
     * @param xmlNodePathColumnIdx Xml path column index
     * @param dataColumnIdx        Xml node data column
     * @param rootXmlBuilderMap    Xml root node map
     * @throws IOException
     */
    private static Map<String, String> initXmlBuilderAndCollectDataByEasyExcel(File excelFile,
                                                                               int xmlNodePathColumnIdx,
                                                                               int dataColumnIdx,
                                                                               Map<String, XMLBuilder2> rootXmlBuilderMap) {


        ExcelReaderBuilder readerBuilder = EasyExcel.read(excelFile);

        System.out.println("Start init xml builder data by easy excel");

        long startTime = System.currentTimeMillis();

        // Collect each row data, key is node path
        Map<String, String> dataMap = Maps.newHashMap();

        List<Map<Integer, String>> datas = readerBuilder.sheet(0).headRowNumber(0).doReadSync();

        for (Map<Integer, String> rowDataMap : datas) {

            // Cell value
            String cellValue;

            String dataKey = null;
            String data = null;

            for (Map.Entry<Integer, String> cellDataEntry : rowDataMap.entrySet()) {

                cellValue = cellDataEntry.getValue();
                cellValue = StringUtils.isBlank(cellValue) ? "" : cellValue.trim();

                // Collect root names from first column
                if (cellDataEntry.getKey() == (xmlNodePathColumnIdx - 1)) {
                    cellValue = cellValue.replaceAll("/+", "/");
                    dataKey = cellValue;
                    String[] slashArrays = StringUtils.split(cellValue, "/");
                    if (slashArrays.length > 0 && !rootXmlBuilderMap.containsKey(slashArrays[0])) {
                        rootXmlBuilderMap.put(slashArrays[0], XMLBuilder2.create(slashArrays[0]));
                    }
                }

                if (cellDataEntry.getKey() == (dataColumnIdx - 1)) {
                    data = cellValue;
                }
            }

            if (null != dataKey) {
                if (dataMap.containsKey(dataKey)) {
                    System.out.println("Already exist xml path: " + dataKey + ", data: " + data + ", will replace with data: " + data);
                }
                dataMap.put(dataKey, data);
            }
        }

        System.out.println("Finish init xml builder data, cost: " + (System.currentTimeMillis() - startTime) + "(ms)");

        return dataMap;
    }

    /**
     * Convert {@link XMLBuilder2} to file
     *
     * @param xmlBuilder2      {@link XMLBuilder2}
     * @param fileRelativePath File relative path
     * @param fileName         File name
     * @throws IOException
     * @throws TransformerException
     */
    private static void xmlBuilderToFile(XMLBuilder2 xmlBuilder2, String fileRelativePath, String fileName) throws IOException, TransformerException {

        long startTime = System.currentTimeMillis();

        Document xmlDocument = xmlBuilder2.getDocument();

        String xmlFileDetailPath = fileRelativePath + "/" + fileName + "-" + System.currentTimeMillis() + ".xml";

        DOMSource source = new DOMSource(xmlDocument);
        FileWriter writer = new FileWriter(xmlFileDetailPath);
        StreamResult result = new StreamResult(writer);

        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();

        // Format xml file
        transformer.setOutputProperty(OutputKeys.ENCODING, StandardCharsets.UTF_8.name());
        transformer.setOutputProperty(OutputKeys.METHOD, "xml");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

        transformer.transform(source, result);

        System.out.println("Generate xml file: [" + xmlFileDetailPath + "] finish, cost: " + (System.currentTimeMillis() - startTime) + "(ms)");
    }

    /**
     * Get excel cell data as string
     *
     * @param cell {@link Cell}
     * @return Data of {@link String}
     */
    private static String getCellValue(Cell cell) {

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        switch (cell.getCellType()) {
            case BOOLEAN:
            case STRING:
            case FORMULA:
                return dataFormatter.formatCellValue(cell);
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    throw new UnsupportedOperationException("Not support date");
                }
                return dataFormatter.formatCellValue(cell);
            default:
                return "";
        }
    }

    /**
     * Check xml node if exist
     *
     * @param original
     * @param nodeName
     * @return
     */
    private static boolean checkIfExist(XMLBuilder2 original, String nodeName) {
        try {
            original.xpathFind("//" + nodeName);
            return true;
        } catch (Throwable t) {
            return false;
        }
    }
}

@Data
class XmlData {
    private String attributeName;
    private String attributeValue;
    private String data;
    private String path;
}