package org.basic.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

/**
 * This class reads data from EXCEL (only .xls format) and generates DELETE scripts.
 * Provide data in EXCEL like a database table.
 *@author ah0666039
 */
public class DeleteScriptGenerator {
	public static final String inputFilePath = "C:/Users/ah0666039/Documents/My Received Files/evnt_rsn_cd_AON.xls";
	public static final String outputFolderPath = "C:/my360Workspaces/360Jul15/ScriptGenerator/SQLScript/LeaveReason/";
	public static final String sysdate = "sysdate";
	public static final String nNull = "null";
	public static final String n_Null = "(null)";
	public static final String systimestamp = "systimestamp";
	public static final String sep_pipe = "|";
	public static final String sep_space = " ";
	public static final String sep_comma = ",";
	public static final String sep_sin_inv_comma = "'";
	public static final String sep_sin_equalto = "=";
	
	/**
	 * This method creates DELETE script.
	 * @param args
	 */
	public static void main(String[] args) {
	 try {
		 	Scanner scanner = new Scanner(System.in);
		 	System.out.println("Enter table name : ");
		 	String tblName = scanner.nextLine().trim().toUpperCase();
		 	if(!(tblName != null) || !(tblName.isEmpty())) {
			 	System.out.println("Enter Column name to match constraints while update, separated by space");
			 	String onColumnCheck = scanner.nextLine().trim();
			 	String[] columnChkArr = new String[5];
			 	columnChkArr = onColumnCheck.split(" ");
				List<String> delCondList = Arrays.asList(columnChkArr);
			 	File fileObj = new File(inputFilePath);
		        FileInputStream file = new FileInputStream(fileObj);
		        HSSFWorkbook workbook = new HSSFWorkbook(file);
		        HSSFSheet sheet = null;
		        for(int i = 0; i < workbook.getNumberOfSheets(); i++) {
		        	if(tblName.equalsIgnoreCase(workbook.getSheetName(i))){
//		        		Get first/desired sheet from the workbook
		        		sheet = workbook.getSheetAt(i);
		        		break;
		        	}
		        }
		        
		        if(sheet != null) {
			        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//		        		Iterate through each rows one by one
			        List<StringBuilder> deleteScript =  new ArrayList<StringBuilder>(0);
			        Iterator<Row> rowIterator = sheet.iterator();
			        List<String> fieldName = new ArrayList<String>(0);
			        while(rowIterator.hasNext()){
			        	Row row = rowIterator.next();
			            if(row.getRowNum() == 0) {
			            	Iterator<Cell> cellIterator = row.cellIterator();
			            	while (cellIterator.hasNext())
			            	{
			            		Cell cell = cellIterator.next();
			            		fieldName.add(cell.getStringCellValue().trim());
			            	}
			            }
			            break;
			        }
			        boolean deleteScriptFlag = false;
			        while (rowIterator.hasNext()) {
			            Row row = rowIterator.next();
		            	Iterator<Cell> cellIteratorCheck = row.cellIterator();
		            	boolean skipEmptyRowFlag = false;
		            	while (cellIteratorCheck.hasNext())
		            	{
		            		Cell cell = cellIteratorCheck.next();
		            		if(cell != null && !(cell.toString().isEmpty())) {
		            			skipEmptyRowFlag = true;
		            			break;
		            		}
		            	}
		            	if(skipEmptyRowFlag){
		            	Iterator<Cell> cellIterator = row.cellIterator();
		            	StringBuilder rowDeleteObj = new StringBuilder();
		            	rowDeleteObj.append("DELETE FROM NUCLEUS." + tblName + " WHERE \n");
		            	int count = 0;
		            	while (cellIterator.hasNext()) {
		            		Cell cell = cellIterator.next();
		            		String  cellVal = cell.toString();
		            		if(cellVal.isEmpty()) {
		            		}else {
		            			boolean delFlag = true;
		            			if(delCondList != null && delCondList.contains(fieldName.get(count))) {
		            				delFlag = false;
		            			}
			            		switch (evaluator.evaluateInCell(cell).getCellType()){
				                    case Cell.CELL_TYPE_NUMERIC:
				                    	int number = (int) cell.getNumericCellValue();
				                    	rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
				                    	rowDeleteObj.append(String.valueOf(number).trim() + " AND \n");
				                        break;
				                    case Cell.CELL_TYPE_STRING:
				                    	if(sysdate.equalsIgnoreCase(cell.getStringCellValue())
				                    		|| n_Null.equalsIgnoreCase(cell.getStringCellValue())
				                    		|| nNull.equalsIgnoreCase(cell.getStringCellValue())
				                    		|| systimestamp.equalsIgnoreCase(cell.getStringCellValue())) {
				                    		if(!delFlag) {
				                    			rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
				                    			rowDeleteObj.append(cell.getStringCellValue().toString().toUpperCase().trim() + " AND \n");
				                    		}
				                    	}else {
				                    		if(!delFlag) {
					                    		rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
					                    		rowDeleteObj.append("'" + cell.getStringCellValue().trim() + sep_sin_inv_comma + " AND \n");
				                    		}
				                    	}
				                        break;
				                    case Cell.CELL_TYPE_BOOLEAN:
				                    	Boolean val = new Boolean(cell.getBooleanCellValue());
				                    		if(!delFlag) {
					                    		rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
					                    		rowDeleteObj.append("'" + val.toString().toUpperCase().trim() + sep_sin_inv_comma + " AND \n");
				                    		}
				                    	break;
				                    case Cell.CELL_TYPE_FORMULA:
				                        break;
			            		}
		            		}
		            		count++;
		            	}
		            	rowDeleteObj.replace(rowDeleteObj.toString().trim().length()-4,rowDeleteObj.length(),"; \n");
		            	deleteScript.add(rowDeleteObj);
		            	if(deleteScript != null && !deleteScript.isEmpty()) {
		            		deleteScriptFlag = true;
		            	}
			        	}
			         }
			        if(deleteScriptFlag) {
				        PrintWriter deleteWriter = new PrintWriter(outputFolderPath  + "DELETE_" + tblName + "_DML.sql", "UTF-8");
				        for(StringBuilder sbObj : deleteScript){
				        	deleteWriter.println(sbObj.toString());
				        }
				        deleteWriter.close();
				        file.close();
				        System.out.println("Delete Scripts generated Successfully");
			        }
			        else {
			        	System.out.println("No data in excel");
			        }
		        }
		        else {
		        	System.out.println("Table name and sheet name in excel are not same");
		        }
		    }else {
		    	System.out.println("Table name cannot be Null or empty");
		    }
	    }
	    catch (Exception e){
	        System.out.println("Exception occured  \n"+ e.getMessage());
	    }
	}
}
