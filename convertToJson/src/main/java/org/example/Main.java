package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class Main {
    public static void main(String[] args) {
        String excelFilePath = "AYNA-data.xlsx";

        readExcelAndWriteToFiles(excelFilePath);
//        String jsonFilePath = "general.json";

//        JSONArray jsonArray = readExcelToJsonArray(excelFilePath);
//        JSONArray jsonArray = readExcelToJsonArray2(excelFilePath);

//        System.out.println(jsonArray);


//        writeJsonArrayToFile(jsonArray, jsonFilePath);
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
                    return df.format(cell.getDateCellValue());
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (String.valueOf(numericValue).contains(".") && numericValue == Math.floor(numericValue)) {
                        return String.valueOf((int) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static String getHeaderName(Sheet sheet, int columnIndex) {
        Row headerRow = sheet.getRow(0);
        Cell headerCell = headerRow.getCell(columnIndex);
        return headerCell.getStringCellValue();
    }

    private static JSONArray readSheetToJSONArray(Sheet sheet) {
        JSONArray jsonArray = new JSONArray();

        Iterator<Row> iterator = sheet.iterator();
        if (iterator.hasNext()) {
            iterator.next();
        }

        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            JSONObject obj = new JSONObject();

            Iterator<Cell> cellIterator = currentRow.iterator();
            int cellIndex = 0;
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();

                String header = getHeaderName(sheet, cellIndex);
                String value = getCellValueAsString(currentCell);

                obj.put(header, value);
                cellIndex++;
            }
            jsonArray.put(obj);
        }

        return jsonArray;
    }

    private static void writeJSONArrayToFile(JSONArray jsonArray, String filePath) {
        try (FileWriter fileWriter = new FileWriter(filePath)) {
            fileWriter.write(jsonArray.toString(4));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void readExcelAndWriteToFiles(String excelFilePath) {
        try (InputStream inputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                JSONArray jsonArray = readSheetToJSONArray(sheet);

                String sheetName = sheet.getSheetName();
                String outputFilePath = excelFilePath.replace(".xlsx", "_" + sheetName + ".json");
                writeJSONArrayToFile(jsonArray, outputFilePath);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}