package pack;
import java.io.File; 
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

public class RetrivingDataInJsonFormat {
	 private JsonObject readExcelFileAsJsonObject_RowWise(String filePath) {
	        DataFormatter dataFormatter = new DataFormatter();
	        JsonObject workbookJson = new JsonObject();
	        JsonArray sheetJson = new JsonArray();
	        JsonObject rowJson = new JsonObject();
	        try {

	            FileInputStream excelFile = new FileInputStream(new File(filePath));
	            Workbook workbook = new XSSFWorkbook(excelFile);
	            FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
	      for (org.apache.poi.ss.usermodel.Sheet sheet : workbook) {
	                sheetJson = new JsonArray();
	                int lastRowNum = ((XSSFSheet) sheet).getLastRowNum();
	                int lastColumnNum = sheet.getRow(0).getLastCellNum();
	                Row firstRowAsKeys = sheet.getRow(0); // first row as a json keys

	                for (int i = 1; i <= lastRowNum; i++) {
	                    rowJson = new JsonObject();
	                    Row row = sheet.getRow(i);
	                    if (row != null) {
	                    	
	                for (int j = 0; j < lastColumnNum; j++) {
	      formulaEvaluator.evaluate(row.getCell(j));
	rowJson.addProperty(firstRowAsKeys.getCell(j).getStringCellValue(),
	 dataFormatter.formatCellValue(row.getCell(j),formulaEvaluator));}
	                        sheetJson.add(rowJson);
	                       }
	                     }
	                workbookJson.add(sheet.getSheetName(), sheetJson); }
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	        return workbookJson;
	    }

	    public static void main(String arg[]) {
	    	RetrivingDataInJsonFormat excelConvertor = new RetrivingDataInJsonFormat();
	        String filePath = ".\\DataFolder\\MFS_CustomerList_05182022.xlsx";
	        JsonObject data = excelConvertor.readExcelFileAsJsonObject_RowWise(filePath);
	        System.out.println(data);
	    }
	  }