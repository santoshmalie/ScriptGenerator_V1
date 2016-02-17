package org.basic.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

/**
 * This method creates Insert scripts.
 * Note : Provide table structure in required format(in .xls file format.) and Name of the file should be same as Table name  
 * @author AH0666039
 */
public class InsertScriptGenerator {

	public static void main(String[] args) {
		 try
		    {
			 	/*Scanner scanner = new Scanner(System.in);
			 	String tblName = "";
			 	System.out.println("Enter table name :");
			 	tblName = scanner.nextLine();*/
			 	File fileObj = new File("C:/Users/ah0666039/Desktop/CICscripts/CodeInserts/SP_CODE_DESCR_TBL.xls");
		        FileInputStream file = new FileInputStream(fileObj);
		        String fileName = FilenameUtils.removeExtension(fileObj.getName());
		        //Create Workbook instance holding reference to .xlsx file
		        HSSFWorkbook workbook = new HSSFWorkbook(file);
		        
//		        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		         
		        //Get first/desired sheet from the workbook
		        HSSFSheet sheet = workbook.getSheetAt(0);
		 
		        //Iterate through each rows one by one
		        final StringBuffer INSERT_SCRIPT = new StringBuffer("INSERT INTO Nucleus."+ fileName + " (");
		        List<StringBuffer> script =  new ArrayList<StringBuffer>(0);
		        Iterator<Row> rowIterator = sheet.iterator();
		        StringBuffer headerRow = new StringBuffer();
		        while(rowIterator.hasNext()){
		        	Row row = rowIterator.next();
		        	headerRow.append(INSERT_SCRIPT);
		            if(row.getRowNum() == 0) {
		            	Iterator<Cell> cellIterator = row.cellIterator();
		            	while (cellIterator.hasNext())
		            	{
		            		Cell cell = cellIterator.next();
		            		headerRow.append(cell.getStringCellValue() + ",");
		            	}
		            	headerRow.replace(headerRow.length()-1, headerRow.length(), "");
		            	headerRow.append(") VALUES (");
		            }
		            break;
		        }
		        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		        
		        while (rowIterator.hasNext())
		        {
		            Row row = rowIterator.next();
	            	Iterator<Cell> cellIterator = row.cellIterator();
	            	StringBuffer rowInsertObj = new StringBuffer();
	            	rowInsertObj.append(headerRow);
	            	while (cellIterator.hasNext())
	            	{
	            		Cell cell = cellIterator.next();
	            		
	            		switch (evaluator.evaluateInCell(cell).getCellType()){
	                    case Cell.CELL_TYPE_NUMERIC:
	                    	int number = (int) cell.getNumericCellValue();
	                    	rowInsertObj.append("" + String.valueOf(number).trim() + " ");
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                    	if("sysdate".equalsIgnoreCase(cell.getStringCellValue())
	                    		|| "null".equalsIgnoreCase(cell.getStringCellValue())) {
	                    		rowInsertObj.append("" +cell.getStringCellValue().toString().toUpperCase().trim() + " ");
	                    	}else {
	                    		rowInsertObj.append("'" + cell.getStringCellValue().trim() + "' ");
	                    	}
	                        break;
	                    case Cell.CELL_TYPE_BOOLEAN:
	                    	Boolean val = new Boolean(cell.getBooleanCellValue());
	                    		rowInsertObj.append("'" + val.toString().toUpperCase().trim() + "' ");
	                    	break;
	                    case Cell.CELL_TYPE_FORMULA:
	                        //Not again
	                        break;
	            		}
	            		
	            		rowInsertObj.append(","); 
	            	}
	            	rowInsertObj.append(");");
	            	rowInsertObj.replace(rowInsertObj.length()-4,rowInsertObj.length(),");");
	            	script.add(rowInsertObj);
	            	System.out.println("");
		         }
		             
		        PrintWriter writer = new PrintWriter("C:/Users/ah0666039/Desktop/CICscripts/CodeInserts/" + fileName + "INSERT.sql", "UTF-8");
		        for(StringBuffer sbObj : script)
		        {
		        	writer.println(sbObj.toString());
		        }
		        writer.close();
		        file.close();
		    }
		    catch (Exception e)
		    {
		        e.printStackTrace();
		    }
	}

}
