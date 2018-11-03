package com.utility.jsonObject;

/**
 * @author Mukundan Kannan
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONObject;
import org.json.XML;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class CommonUtil {

    private static final Logger LOGGER = Logger.getLogger(CommonUtil.class);

    public static JsonObject getExcelDataAsJsonObject(File excelFile) throws InvalidFormatException {
        JsonObject sheetsJsonObject = new JsonObject();
        Workbook workbook = null;
        try {
            // workbook = new XSSFWorkbook(excelFile);
            workbook = WorkbookFactory.create(excelFile);
        } catch (IOException e) {
            e.printStackTrace();
            LOGGER.error(e.getMessage(), e);
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            JsonArray sheetArray = new JsonArray();
            ArrayList<String> columnNames = new ArrayList<String>();
            Sheet sheet = workbook.getSheetAt(i);
            Iterator<Row> sheetIterator = sheet.iterator();
            // iterating the sheet
            while (sheetIterator.hasNext()) {
                Row currentRow = sheetIterator.next();
                JsonObject jsonObject = new JsonObject();
                // getting row count and values and adding to the object
                if (currentRow.getRowNum() != 0) {
                    for (int j = 0; j < columnNames.size(); j++) {
                        if (currentRow.getCell(j) != null) {
                            if (currentRow.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {
                                jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getStringCellValue());
                            } else if (currentRow.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getNumericCellValue());
                            } else if (currentRow.getCell(j).getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                                jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getBooleanCellValue());
                            } else if (currentRow.getCell(j).getCellType() == Cell.CELL_TYPE_BLANK) {
                                jsonObject.addProperty(columnNames.get(j), "");
                            }
                        } else {
                            jsonObject.addProperty(columnNames.get(j), "");
                        }
                    }
                    sheetArray.add(jsonObject);
                } else {
                    // getting column name and storing it
                    for (int k = 0; k < currentRow.getPhysicalNumberOfCells(); k++) {
                        columnNames.add(currentRow.getCell(k).getStringCellValue());
                    }
                }

            }
            sheetsJsonObject.add(workbook.getSheetName(i), sheetArray);
        }
        return sheetsJsonObject;
    }

    public static JSONObject getXmlDataAsJsonObject(File file) {
        JSONObject jsonObj = new JSONObject();
        try {
            InputStream inputStream = new FileInputStream(file);
            StringBuilder builder = new StringBuilder();
            int ptr = 0;
            while ((ptr = inputStream.read()) != -1) {
                builder.append((char) ptr);
            }
            String xml = builder.toString();
            jsonObj = XML.toJSONObject(xml);
            inputStream.close();
        }
        catch (Exception e) {
            e.printStackTrace();
            LOGGER.error(e.getMessage(), e);
        }
        return jsonObj;
    }

    /**
     * Convert a JSON string to pretty print version
     * @param jsonString
     * @return
     */
    public static String toPrettyFormat(String jsonString) {
        JsonParser parser = new JsonParser();
        JsonObject json = parser.parse(jsonString).getAsJsonObject();

        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        String prettyJson = gson.toJson(json);

        return prettyJson;
    }


    /*
     * public static void main(String args[]) throws InvalidFormatException {
     * // String fileLocation = "C://Homedepot//All 300 Collection Title Description and curation ID.xlsx";
     * String fileLocation = "C://Homedepot//MCB-Eclipse-Style.xml";
     * 
     * File newFile = new File(fileLocation);
     * // JsonObject getJson = CommonUtil.getExcelDataAsJsonObject(newFile);
     * JSONObject getJson = CommonUtil.getXmlDataAsJsonObject(newFile);
     * LOGGER.info("Json object :" + getJson);
     * }
     */

}
