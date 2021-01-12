package org.jobchen.training;

import com.google.common.collect.Maps;
import com.jamesmurty.utils.XMLBuilder2;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.w3c.dom.Document;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Map;

/**
 * @author : jobchen
 * @version : 1.0.0 Copyright (c) 2021 jobchen, All Rights Reserved.
 * @program : Training
 * @description : Parse excel data to xml
 * @create : 2021-01-12 55:29
 * @history : Who--->When--->What
 */
public class ParseExcelToXml {

    public static void main(String[] args) throws IOException, TransformerException {
        parse("/Users/jobchen/Desktop/sample001.xlsx", 1, 4, "/Users/jobchen/Desktop");

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
     * @param excelPath            excel file path
     * @param xmlNodePathColumnIdx xml node path index, start from {@code 1}
     * @param dataColumnIdx        data index, start from {@code 1}
     * @param xmlFilePath          generated xml file path, this is file path not include file name, file name is decided by xml root
     * @throws IOException
     * @throws TransformerException
     */
    private static void parse(String excelPath,
                              int xmlNodePathColumnIdx,
                              int dataColumnIdx,
                              String xmlFilePath) throws IOException, TransformerException {

        Workbook workbook = WorkbookFactory.create(new File(excelPath));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " sheets : ");

        // Handle first sheet data
        Sheet sheet = workbook.getSheetAt(0);

        // Collect root node name and a init root xml builder
        Map<String, XMLBuilder2> rootXmlBuilderMap = Maps.newHashMap();

        // Collect each row data, key is node path
        Map<String, String> dataMap = Maps.newHashMap();

        initXmlBuilderData(new File(excelPath), xmlNodePathColumnIdx, dataColumnIdx, rootXmlBuilderMap, dataMap);

        if (MapUtils.isNotEmpty(dataMap)) {

            String[] dataKeys;

            for (Map.Entry<String, String> dataEntry : dataMap.entrySet()) {

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
                                rootXmlBuilderMap.get(dataKeys[0]).xpathFind("//" + dataKeys[i - 1]).e(dataKeys[i]);
                            }
                        } else {
                            rootXmlBuilderMap.get(dataKeys[0]).xpathFind("//" + dataKeys[i - 1]).e(dataKeys[i]).t(dataEntry.getValue());
                        }
                    }
                }

                System.out.println(dataEntry.getKey() + ": " + rootXmlBuilderMap.get(dataKeys[0]).asString());
            }

            for (Map.Entry<String, XMLBuilder2> xmlBuilder2Entry : rootXmlBuilderMap.entrySet()) {
                xmlBuilderToFile(xmlBuilder2Entry.getValue(), xmlFilePath, xmlBuilder2Entry.getKey());
            }

        }

        // Closing the workbook
        workbook.close();
    }

    /**
     * Build xml builder precondition data
     *
     * @param excelFile            Excel file
     * @param xmlNodePathColumnIdx Xml path column index
     * @param dataColumnIdx        Xml node data column
     * @param rootXmlBuilderMap    Xml root node map
     * @param dataMap              Xml data map
     * @throws IOException
     */
    private static void initXmlBuilderData(File excelFile,
                                           int xmlNodePathColumnIdx,
                                           int dataColumnIdx,
                                           Map<String, XMLBuilder2> rootXmlBuilderMap,
                                           Map<String, String> dataMap) throws IOException {

        Workbook workbook = WorkbookFactory.create(excelFile);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " sheets : ");

        // Handle first sheet data
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {

            // Cell value
            String cellValue;

            String dataKey = null;
            String data = null;

            // Record column of a row
            int cellIndex = 0;

            for (Cell cell : row) {

                cellValue = getCellValue(cell);

                // Collect root names from first column
                if (cellIndex == (xmlNodePathColumnIdx - 1)) {
                    cellValue = cellValue.trim();
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

        Document xmlDocument = xmlBuilder2.getDocument();

        String xmlFileDetailPath = fileRelativePath + "/" + fileName + "-" + System.currentTimeMillis() + ".xml";

        DOMSource source = new DOMSource(xmlDocument);
        FileWriter writer = new FileWriter(xmlFileDetailPath);
        StreamResult result = new StreamResult(writer);

        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.transform(source, result);

        System.out.println("Generate xml file: [" + xmlFileDetailPath + "] finish");
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
