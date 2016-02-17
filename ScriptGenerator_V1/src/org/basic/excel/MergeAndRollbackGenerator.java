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
 * This class reads data from EXCEL (only .xls format) and creates MERGE and DELETE scripts 
 * in two different SQL files.
 * Input : First provide table name, then provide name of the columns,separated by pipe, to satisfy table constraints.
 * Note  : If unique key contains all the columns then please comment code delimited by (NO_UPDATE_REQUIRED)
 * Provide data in EXCEL like a database table.
 *@author ah0666039
 */
public class MergeAndRollbackGenerator {
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
	 * This method creates merge script.
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
			        final StringBuffer MERGE_SCRIPT = new StringBuffer("MERGE INTO NUCLEUS." + tblName + " MOBJ \n USING ( \n ");
			        List<StringBuilder> insertScript =  new ArrayList<StringBuilder>(0);
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
			        boolean createScriptFlag = false;
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
		            	StringBuilder rowInsertObj = new StringBuilder();
		            	StringBuilder rowDeleteObj = new StringBuilder();
		            	rowDeleteObj.append("DELETE FROM NUCLEUS." + tblName + " WHERE \n");
		            	rowInsertObj.append(MERGE_SCRIPT);
		            	rowInsertObj.append("SELECT");
		            	int count = 0;
		            	while (cellIterator.hasNext()) {
		            		Cell cell = cellIterator.next();
		            		String  cellVal = cell.toString();
		            		if(cellVal.isEmpty()) {
		            			rowInsertObj.append("\n \t '' ");
		            		}else {
		            			boolean delSkipFlag = false;
		            			if(delCondList != null && delCondList.contains(fieldName.get(count))) {
		            				delSkipFlag = true;
		            			}
			            		switch (evaluator.evaluateInCell(cell).getCellType()){
				                    case Cell.CELL_TYPE_NUMERIC:
				                    	int number = (int) cell.getNumericCellValue();
				                    	rowInsertObj.append("\n \t" + String.valueOf(number).trim() + sep_space);
				                    	rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
				                    	rowDeleteObj.append(String.valueOf(number).trim() + " AND \n");
				                        break;
				                    case Cell.CELL_TYPE_STRING:
				                    	if(sysdate.equalsIgnoreCase(cell.getStringCellValue())
				                    		|| n_Null.equalsIgnoreCase(cell.getStringCellValue())
				                    		|| nNull.equalsIgnoreCase(cell.getStringCellValue())
				                    		|| systimestamp.equalsIgnoreCase(cell.getStringCellValue())) {
				                    		rowInsertObj.append("\n \t" +cell.getStringCellValue().toString().toUpperCase().trim() + sep_space);
				                    		if(!delSkipFlag) {
				                    			rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
				                    			rowDeleteObj.append(cell.getStringCellValue().toString().toUpperCase().trim() + " AND \n");
				                    		}
				                    	}else {
				                    		rowInsertObj.append("\n \t'" + cell.getStringCellValue().trim() + sep_sin_inv_comma + sep_space);
				                    		if(!delSkipFlag) {
					                    		rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
					                    		rowDeleteObj.append("'" + cell.getStringCellValue().trim() + sep_sin_inv_comma + " AND \n");
				                    		}
				                    	}
				                        break;
				                    case Cell.CELL_TYPE_BOOLEAN:
				                    	Boolean val = new Boolean(cell.getBooleanCellValue());
				                    		rowInsertObj.append("\n \t'" + val.toString().toUpperCase().trim() + sep_sin_inv_comma + sep_space);
				                    		if(!delSkipFlag) {
					                    		rowDeleteObj.append("\t" + fieldName.get(count) + sep_space + sep_sin_equalto + sep_space);
					                    		rowDeleteObj.append("'" + val.toString().toUpperCase().trim() + sep_sin_inv_comma + " AND \n");
				                    		}
				                    	break;
				                    case Cell.CELL_TYPE_FORMULA:
				                        break;
			            		}
		            		}
		            		rowInsertObj.append(fieldName.get(count) + sep_comma + sep_space);
		            		count++;
		            	}
		            	rowInsertObj.replace(rowInsertObj.length()-2,rowInsertObj.length()," FROM DUAL )  ROBJ \n");
		            	rowInsertObj.append("ON (");
		            	count = 0;
		            	for(int i = 0; i < columnChkArr.length; i++ ) {
		            		StringBuffer str = new StringBuffer(columnChkArr[i]);
		            		if(str != null && !(str.length() == 0)) {
		            			rowInsertObj.append(" \n MOBJ." + str.toString().trim() + " = ROBJ." + str.toString().trim() + " AND ");
		            		}
		            	}
		            	rowInsertObj.replace(rowInsertObj.length()-5,rowInsertObj.length()," ) \n");
//		            	NO_UPDATE_REQUIRED
		            	rowInsertObj.append("WHEN MATCHED THEN UPDATE SET \n");
		            	count = 0;
		            	for(String field : fieldName) {
		            		if(field.equalsIgnoreCase(fieldName.get(0))){
		            			continue;
		            		}else {
		            			List<String> colList = Arrays.asList(columnChkArr);
		            			if(colList.contains(field.trim())) {
		            				continue;
		            			}
		            			rowInsertObj.append("MOBJ." + field.toString().trim() + " = " + "ROBJ." + field.toString().trim() + sep_comma);
		            			if(count > 3) {
		            				rowInsertObj.append("\n");
		            				count = 0;
		            			}
		            		}
		            		count++;
		            	}
		            	rowInsertObj.replace(rowInsertObj.toString().trim().length()-1,rowInsertObj.length(),"\n");
//		            	NO_UPDATE_REQUIRED
		            	rowInsertObj.replace(rowInsertObj.length()-1,rowInsertObj.length(),"\n WHEN NOT MATCHED THEN INSERT \n (");
		            	for(String fieldString : fieldName) {
		            		rowInsertObj.append("MOBJ." + fieldString.trim() + sep_comma);
		            	}
		            	rowInsertObj.replace(rowInsertObj.length()-1,rowInsertObj.length(),") \n");
		            	rowInsertObj.append("VALUES \n (");
		            	for(String fieldString : fieldName) {
		            		rowInsertObj.append("ROBJ." + fieldString.trim() + sep_comma);
		            	}
		            	rowInsertObj.replace(rowInsertObj.length()-1,rowInsertObj.length(),"); \n");
		            	rowDeleteObj.replace(rowDeleteObj.toString().trim().length()-4,rowDeleteObj.length(),"; \n");
		            	insertScript.add(rowInsertObj);
		            	deleteScript.add(rowDeleteObj);
		            	if(insertScript != null && !insertScript.isEmpty()) {
		            		createScriptFlag = true;
		            	}
			        	}
			         }
			        if(createScriptFlag) {
				        PrintWriter insertWriter = new PrintWriter(outputFolderPath  + "MERGE_" + tblName + "_DML.sql", "UTF-8");
				        for(StringBuilder sbObj : insertScript){
				        	insertWriter.println(sbObj.toString());
				        }
				        
				        PrintWriter deleteWriter = new PrintWriter(outputFolderPath  + "DELETE_" + tblName + "_DML.sql", "UTF-8");
				        for(StringBuilder sbObj : deleteScript){
				        	deleteWriter.println(sbObj.toString());
				        }
				        
				        insertWriter.close();
				        deleteWriter.close();
				        file.close();
				        System.out.println("Merge Scripts generated Successfully");
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
	    catch (Exception e)
	    {
	        System.out.println("Exception occured  \n"+ e.getMessage());
	    }
	}
}
