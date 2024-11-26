package com.garguir;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    private static final Logger logger = LogManager.getLogger("ExcelToJson");

    public static void main(String[] args) {
        logger.info("Start");
        String excelFilePath = "staff.xlsx";
        List<Map<String, Object>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            List<String> headers = new ArrayList<>();

            // Obtener los encabezados
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            // Leer los datos
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Comenzar desde la segunda fila
                Row row = sheet.getRow(i);
                Map<String, Object> rowData = new HashMap<>();

                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    if(cell == null){
                        rowData.put(headers.get(j)," ");
                    }else {
                        rowData.put(headers.get(j), getCellValue(cell));
                    }
                }
                data.add(rowData);
            }

            // Convertir a JSON
            ObjectMapper objectMapper = new ObjectMapper();
            String jsonOutput = objectMapper.writeValueAsString(data);
            System.out.println(jsonOutput);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    DataFormatter formatter = new DataFormatter();
                    return formatter.formatCellValue(cell);
                }else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            default:
                return null;
        }
    }
}