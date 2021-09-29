package com.lipogramsw.jitools.xlsfiles;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.MonthDay;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lipogramsw.jitools.LogHandler;

public class XLSConverter {

	public XLSConverter() 
	{
	    //LogHandler.getInstance().writeLog("Java Integration Tools for Thomson Reuters - CSV to XLS/XLSX");
	    //LogHandler.getInstance().writeLog("http://www.lipogramsw.com/jitools");
	}
	
	
	public void convertFile(String fileIn, String fileOut, String sheetName, String separator)
	{
		this.internalConvertFile(fileIn, fileOut, sheetName, separator);		
	}
	
	public void convertFileTemplate(String fileIn, String fileOut, String fileTemplate, int startLine, String sheetName, String separator)
	{
		this.internalConvertFileTemplate(fileIn, fileOut, fileTemplate, startLine, sheetName, separator);
	}
	
	// Added 30.04.2021: Overload for accept outputType ('xls' or 'xlsx')
	public void convertFile(String fileIn, String fileOut, String sheetName, String separator, String outputType)
	{
		if (outputType.equalsIgnoreCase("XLSX"))
		{
			this.internalConvertFileXLSX(fileIn, fileOut, sheetName, separator);
		}
		else
		{
			LogHandler.getInstance().writeLog("HINT: XLSX files are better handled by this application.");
			LogHandler.getInstance().writeLog("      XLS actually export numbers/dates as text, not formatted. ");
			this.internalConvertFile(fileIn, fileOut, sheetName, separator);
		}
		
	}
	
	public void convertFileTemplate(String fileIn, String fileOut, String fileTemplate, int startLine, String sheetName, String separator, String outputType)
	{
		if (outputType.equalsIgnoreCase("XLSX"))
		{
			this.internalConvertFileTemplateXLSX(fileIn, fileOut, fileTemplate, startLine, sheetName, separator);
		}
		else
		{
			LogHandler.getInstance().writeLog("HINT: XLSX files are better handled by this application.");
			LogHandler.getInstance().writeLog("      XLS actually export numbers/dates as text, not formatted. ");
			this.internalConvertFileTemplate(fileIn, fileOut, fileTemplate, startLine, sheetName, separator);
		}
	}
	
	private void internalConvertFile(String fileIn, String fileOut, String sheetName, String separator) 
	  {
	    String thisSeparator = ",";
	    
	    if (separator != null)
	    {
	      thisSeparator = separator;
	    }
	    
	    LogHandler.getInstance().writeLog("Field separator in CSV is '" + thisSeparator + "'");
	    LogHandler.getInstance().writeLog("Reading input file '" + fileIn + "'");

	    ArrayList<ArrayList<String>> arList = new ArrayList<>();
	    ArrayList<String> al = null;
	    
	    try 
	    {
	      BufferedReader myInput = new BufferedReader(new InputStreamReader(new FileInputStream(fileIn)));
	      
	      String thisLine;
	      	      
	      while ((thisLine = myInput.readLine()) != null) 
	      {
	        al = new ArrayList<>();
	        String[] strar = thisLine.split(thisSeparator, -1);
	        for (int j = 0; j < strar.length; j++) 
	        {
	          String edit = strar[j].replace('\n', ' ');
	          al.add(edit);
	        } 
	        arList.add(al);
	      } 
	      
	      LogHandler.getInstance().writeLog("Creating new XLS file '" + fileOut + "'");
	      HSSFWorkbook hwb = new HSSFWorkbook();
	      HSSFSheet sheet = null;
	      if (sheetName == null) 
	      {
	        sheet = hwb.createSheet("Sheet1");
	      }
	      else 
	      {
	        sheet = hwb.createSheet(sheetName);
	      } 
	      
	      for (int k = 0; k < arList.size(); k++) 
	      {
	        ArrayList<String> ardata = arList.get(k);
	        HSSFRow row = sheet.createRow(0 + k);
	        
	        for (int p = 0; p < ardata.size(); p++) 
	        {
	          HSSFCell cell = row.createCell((short)p);
	          cell.setCellValue(((String)ardata.get(p)).toString());
	        } 
	      } 
	      
	      myInput.close();
	      
	      LogHandler.getInstance().writeLog("Writing output file '" + fileOut + "' (overwriting if existent)");
	      FileOutputStream fileConverted = new FileOutputStream(fileOut);
	      hwb.write(fileConverted);
	      fileConverted.close();
	      
	      hwb.close();
	      
	      LogHandler.getInstance().writeLog("Finished. ");
	    }
	    catch (Exception ex) 
	    {
	      LogHandler.getInstance().writeLog("ERROR: " + ex.getMessage());
	      ex.printStackTrace();
	      LogHandler.getInstance().writeLog("Execution aborted.");
	      System.exit(4);
	    } 
	  }
	  
	  
	  private void internalConvertFileTemplate(String fileIn, String fileOut, String fileTemplate, int startLine, String sheetName, String separator) 
	  {
		Workbook wbTemplate = null; 
		Sheet templateSheet = null;
		Row templateRow = null;
		Cell templateCell = null;
	    String thisSeparator = ",";
	    
	    if (separator != null)
	    {
	      thisSeparator = separator;
	    }
	    LogHandler.getInstance().writeLog("Field separator in CSV is '" + thisSeparator + "'");
	    LogHandler.getInstance().writeLog("Will use template file '" + fileTemplate + "'");
	    try
	    {
	    	wbTemplate = WorkbookFactory.create(new File(fileTemplate));
	    	LogHandler.getInstance().writeLog("Template file starts at line " + startLine);
	    	templateSheet = wbTemplate.getSheetAt(0);
	    	templateRow = templateSheet.getRow(startLine -1);
	    	LogHandler.getInstance().writeLog(" - Template line has " + templateRow.getLastCellNum() + " columns.");
	    	for (Cell cell : templateRow)
	    	{
	    		if (cell.getCellType() == CellType.FORMULA)
	    			LogHandler.getInstance().writeLog(" - Cell " + (cell.getColumnIndex() + 1) + " contains a formula and will be ignored. ");
	    	}
	    }
	    catch(Exception e)
	    {
	    	LogHandler.getInstance().writeLog("Error loading template file; aborting execution.");
	    	System.exit(4);
	    }
	    
	    LogHandler.getInstance().writeLog("Reading input file '" + fileIn + "'");
	    ArrayList<ArrayList<String>> arList = new ArrayList<>();
	    ArrayList<String> al = null;
	    try 
	    {
	      BufferedReader myInput = new BufferedReader(new InputStreamReader(new FileInputStream(fileIn)));
	      String thisLine;
	      while ((thisLine = myInput.readLine()) != null) 
	      {
	        al = new ArrayList<>();
	        String[] strar = thisLine.split(thisSeparator, -1);
	        for (int j = 0; j < strar.length; j++) 
	        {
	          String edit = strar[j].replace('\n', ' ');
	          al.add(edit);
	        } 
	        arList.add(al);
	      } 
	      
	      LogHandler.getInstance().writeLog("Creating new XLS file '" + fileOut + "', based on template '" + fileTemplate + "'");
	      HSSFWorkbook hwb = (HSSFWorkbook)wbTemplate;
	      HSSFSheet sheet = hwb.getSheetAt(0);
	      if (sheetName != null) 
	      {
	        hwb.setSheetName(0, sheetName);
	      } 
	      
	      for (int k = 0; k < arList.size(); k++) 
	      {
	        ArrayList<String> ardata = arList.get(k);
	        HSSFRow row = sheet.createRow((startLine - 1) + k);
	        
	        for (int p = 0; p < ardata.size(); p++) 
	        {
	          HSSFCell cell = row.createCell((short)p);
	          
	          if (p > (templateRow.getLastCellNum() -1))
	          {
	        	  cell = row.createCell((short)p);
	        	  cell.setCellValue(((String)ardata.get(p)).toString());
	       	  }
	          else
	          {	        	  
	        	  templateCell = templateRow.getCell(p);
		          cell.setCellStyle(templateCell.getCellStyle());
		          if (templateCell.getCellType() == CellType.STRING)
		          {		        	  
		        	  cell.setCellType(templateCell.getCellType());
		        	  cell.setCellValue(((String)ardata.get(p)).toString());
		          }
		          
		          if (templateCell.getCellType() == CellType.NUMERIC)
		          {		        	  
		        	  try {
		        		  cell.setCellType(templateRow.getCell(p).getCellType());
		        		  cell.setCellValue(Double.parseDouble(ardata.get(p)));
		        	  } catch (Exception fmt) {
		        		  // Falls back to String if number is invalid.
			        	  cell.setCellType(CellType.STRING);
			        	  cell.setCellValue(((String)ardata.get(p)).toString());
		        	  }
		          }
		          
		          if (templateCell.getCellType() == CellType.FORMULA)
		          {
		        	  cell.setCellType(CellType.STRING);
		        	  cell.setCellValue("<FORMULA>");
		          }
	          }
	        } 
	      } 
	      
	      myInput.close();
	      
	      LogHandler.getInstance().writeLog("Writing output file '" + fileOut + "' (overwriting if existent)");
	      FileOutputStream fileConverted = new FileOutputStream(fileOut);
	      hwb.write(fileConverted);
	      fileConverted.close();
	      
	      hwb.close();
	      
	      LogHandler.getInstance().writeLog("Finished. ");
	    }
	    catch (Exception ex) 
	    {
	      LogHandler.getInstance(). writeLog("ERROR: " + ex.getMessage());
	      ex.printStackTrace();
	      LogHandler.getInstance().writeLog("Execution aborted.");
	      System.exit(4);
	    } 
	  }
	  
	  
	  // Added 30.04.2021:
	  // Functions to use ApachePOI XLSX instead of XLS
	  private void internalConvertFileXLSX(String fileIn, String fileOut, String sheetName, String separator) 
	  {
	    String thisSeparator = ",";
	    
	    if (separator != null)
	    {
	      thisSeparator = separator;
	    }
	    
	    LogHandler.getInstance().writeLog("Field separator in CSV is '" + thisSeparator + "'");
	    LogHandler.getInstance().writeLog("Reading input file '" + fileIn + "'");

	    ArrayList<ArrayList<String>> arList = new ArrayList<>();
	    ArrayList<String> al = null;
	    
	    try 
	    {
	      BufferedReader myInput = new BufferedReader(new InputStreamReader(new FileInputStream(fileIn)));
	      
	      String thisLine;
	      while ((thisLine = myInput.readLine()) != null) 
	      {
	        al = new ArrayList<>();
	        String[] strar = thisLine.split(thisSeparator, -1);
	        for (int j = 0; j < strar.length; j++) 
	        {
	          String edit = strar[j].replace('\n', ' ');
	          al.add(edit);
	        } 
	        arList.add(al);
	      } 
	      
	      LogHandler.getInstance().writeLog("Creating new XLSX file '" + fileOut + "', without a model.");
	      LogHandler.getInstance().writeLog(" - Be aware that all numbers and dates will be exported as text!");
	      XSSFWorkbook hwb = new XSSFWorkbook();
	      XSSFSheet sheet = null;
	      if (sheetName == null) 
	      {
	        sheet = hwb.createSheet("Sheet1");
	      }
	      else 
	      {
	        sheet = hwb.createSheet(sheetName);
	      } 
	      
	      for (int k = 0; k < arList.size(); k++) 
	      {
	        ArrayList<String> ardata = arList.get(k);
	        XSSFRow row = sheet.createRow(0 + k);
	        
	        for (int p = 0; p < ardata.size(); p++) 
	        {
	          XSSFCell cell = row.createCell((short)p);
	          cell.setCellValue(((String)ardata.get(p)).toString());
	        } 
	      } 
	      
	      myInput.close();
	      
	      LogHandler.getInstance().writeLog("Writing output file '" + fileOut + "' (overwriting if existent)");
	      FileOutputStream fileConverted = new FileOutputStream(fileOut);
	      hwb.write(fileConverted);
	      fileConverted.close();
	      
	      hwb.close();
	      
	      LogHandler.getInstance().writeLog("Finished. ");
	    }
	    catch (Exception ex) 
	    {
	      LogHandler.getInstance().writeLog("ERROR: " + ex.getMessage());
	      ex.printStackTrace();
	      LogHandler.getInstance().writeLog("Execution aborted.");
	      System.exit(4);
	    } 
	  }
	  
	  private void internalConvertFileTemplateXLSX(String fileIn, String fileOut, String fileTemplate, int startLine, String sheetName, String separator) 
	  {		
		Workbook wbTemplate = null; 
		Sheet templateSheet = null;
		Row templateRow = null;
		Cell templateCell = null;
	    String thisSeparator = ",";
	    
	    
	    
	    if (separator != null)
	    {
	      thisSeparator = separator;
	    }
	    LogHandler.getInstance().writeLog("Field separator in CSV is '" + thisSeparator + "'");
	    LogHandler.getInstance().writeLog("Will use template file '" + fileTemplate + "'");
	    try
	    {
	    	wbTemplate = WorkbookFactory.create(new File(fileTemplate));
	    	LogHandler.getInstance().writeLog("Template file starts at line " + startLine);
	    	templateSheet = wbTemplate.getSheetAt(0);
	    	templateRow = templateSheet.getRow(startLine -1);
	    	LogHandler.getInstance().writeLog(" - Template line has " + templateRow.getLastCellNum() + " columns.");
	    	for (Cell cell : templateRow)
	    	{
	    		if (cell.getCellType() == CellType.FORMULA)
	    			LogHandler.getInstance().writeLog(" - Cell " + (cell.getColumnIndex() + 1) + " contains a formula and will be ignored. ");
	    	}
	    }
	    catch(Exception e)
	    {
	    	LogHandler.getInstance().writeLog("Error loading template file; aborting execution.");
	    	System.exit(4);
	    }
	    
	    LogHandler.getInstance().writeLog("Reading input file '" + fileIn + "'");
	    ArrayList<ArrayList<String>> arList = new ArrayList<>();
	    ArrayList<String> al = null;
	    try 
	    {
	      BufferedReader myInput = new BufferedReader(new InputStreamReader(new FileInputStream(fileIn)));
	      
	      String thisLine;
	      while ((thisLine = myInput.readLine()) != null) 
	      {
	        al = new ArrayList<>();
	        String[] strar = thisLine.split(thisSeparator, -1);
	        for (int j = 0; j < strar.length; j++) 
	        {
	          String edit = strar[j].replace('\n', ' ');
	          al.add(edit);
	        } 
	        arList.add(al);	        
	      } 
	      
	      LogHandler.getInstance().writeLog("Creating new XLSX file '" + fileOut + "', based on template '" + fileTemplate + "'");
	      InputStream templateStream = new FileInputStream(fileTemplate);					
	      XSSFWorkbook hwb = (XSSFWorkbook) WorkbookFactory.create(templateStream);
	      XSSFSheet sheet = hwb.getSheetAt(0);
	      if (sheetName != null) 
	      {
	        hwb.setSheetName(0, sheetName);
	      } 
	      
	      // Creates an array of cell styles based on the template row
	      // Done this way because the process is too slow; can't create a new style for each cell in the spreadsheet. 
	      // This way, a single set is created and then reused through the file.
	      ArrayList<CellStyle> templateStyle = new ArrayList<>();
	      for (Cell cell : templateRow)
	    	{
    			CellStyle templateCellStyle = hwb.createCellStyle();				// Create new object
    			templateCellStyle.cloneStyleFrom(cell.getCellStyle());				// Copy from template
    			templateStyle.add(templateCellStyle);								// Add to template array
	    	}	      
	      
	      // Iterates through data
	      for (int k = 0; k < arList.size(); k++) 
	      {
	        ArrayList<String> ardata = arList.get(k);
        	XSSFRow row = sheet.createRow((startLine - 1) + k);
        		        
	        for (int p = 0; p < ardata.size(); p++) 
	        {
	          XSSFCell cell = row.createCell((short)p);
	          
	          if (p > (templateRow.getLastCellNum() -1))
	          {
	        	  // Data is larger than template: create new blank / unformatted cell, and
	        	  // sets the data as a simple string format (general)
	        	  cell.setCellValue(((String)ardata.get(p)).toString());
	       	  }
	          else
	          {
	        	  // Data is inside the template area. Format it just like 
	        	  templateCell = templateRow.getCell(p);
	        	  cell.setCellStyle(templateStyle.get(p));							// Sets the cell style/format based on template array

	        	  // String cells are not formatted
	        	  if (templateCell.getCellType() == CellType.STRING)
		          {
	        		  if ((ardata.get(p) == null) || (((String)ardata.get(p)).toString().isEmpty()))
	        		  {
	        			  cell.setBlank();
	        			  //cell.setCellValue(" ");
	        		  }
	        		  else
	        		  {
	        			 cell.setCellValue(((String)ardata.get(p)).toString());
	        		  }
		          }
		          	        	  
	        	  // Numeric cells can be both Numbers or Dates. This is tricky. 
	        	  // Try to convert the numbers to double, but dates are just strings.  
		          if (templateCell.getCellType() == CellType.NUMERIC)
		          {
		        	  boolean localError = false;

		        	  if ((ardata.get(p) == null) || (((String)ardata.get(p)).toString().isEmpty()))
		        	  {
		        		  // NULL value! Set as blank.
		        		  localError = true;
		        		  //cell.setCellValue(" ");
		        		  cell.setBlank();
		        	  }
		        	  else
		        	  {
			        	  // Try converting from to Double (Excel does not use Float)
			        	  // Numeric values should be provided without thousands separator, and with "." as decimal separator.
			        	  try {
			        		  localError = false;
			        		  double localValue = 0;
			        		  localValue = Double.parseDouble(ardata.get(p));
			        		  cell.setCellValue(localValue);
			        	  } catch (Exception fmt) {
			        		  localError = true;
			        	  }
			        	  
			        	  // If resulted in error, replaces "," for "." and try again - just in case...
			        	  try {
			        		  localError = false;
			        		  double localValue = 0;
			        		  localValue = Double.parseDouble(ardata.get(p).replace(',', '.'));
			        		  cell.setCellValue(localValue);
			        	  } catch (Exception fmt) {
			        		  localError = true;
			        	  }
			        	  // Not a number! Fallback to String. Dates *MAY* be handled correctly if defined in template. 
			        	  // Tested formats are YYYY-MM-DD and DD/MM/YYYY. Excel shows them exactly as received, but at least treats 
			        	  // the data like a date cell (can make calculations and stuff).
			        	  if (localError)			        		  
			        	  {
			        		  //Falls back to String, hoping it is a valid date.
				        	  cell.setCellType(CellType.STRING);
				        	  cell.setCellValue(((String)ardata.get(p)).toString());
			        		  
			        		 Matcher matcherDate = Pattern.compile("(.{2}\\/.{2}\\/\\d{2,4})").matcher(ardata.get(p)); //Verifica se é uma Data que está vindo
			        		 while (matcherDate.find()) {
			        			 if(!matcherDate.group(1).isEmpty()) { //Se o group 1 estiver vazio o não é uma data
			        				 SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
			        				 //LogHandler.getInstance().writeLog("Data convertida 1: "+matcherDate.group(1));
					        		 Date date = DateUtils.truncate((Date)formatter.parse(ardata.get(p)),java.util.Calendar.DAY_OF_MONTH);
			        				 date = (Date) formatter.parse(matcherDate.group(1));
					        		 //LogHandler.getInstance().writeLog("Data convertida 2: "+date);
									 cell.setCellValue(date);	
									 //
			        			 }
			        		 }			        	  
			        		  //LogHandler.getInstance().writeLog((String)ardata.get(p));
			        		  //Falls back to String, hoping it is a valid date.
				        	  //cell.setCellType(CellType.STRING);
				        	  //cell.setCellValue(((String)ardata.get(p)).toString());							  
							  //
			        	  }
		        	  }
		          }		         
		          
		          // Nope, we are not going to use formulas here. The problem is not the formula itself,
		          // but the fact that we should treat the line/columns substitutions accordingly. It could 
		          // wreak all sorts of hellish havoc. 
		          if (templateCell.getCellType() == CellType.FORMULA)
		          {
		        	  cell.setCellType(CellType.STRING);
		        	  cell.setCellValue("<FORMULA>");
		          }
	          }
	        } 
	      } 
	      
	      // Closes the input file
	      myInput.close();
	      
	      // Writes the output file. It *will* cause errors if the destination file is already open by another process,
	      // like Excel itself. Don't try to test with the output file open! 
	      LogHandler.getInstance().writeLog("Writing output file '" + fileOut + "' (overwriting if existent)");
	      FileOutputStream fileConverted = new FileOutputStream(fileOut);
	      hwb.write(fileConverted);
	      fileConverted.close();
	      hwb.close();
	      
	      LogHandler.getInstance().writeLog("Finished. ");
	    }
	    catch (Exception ex) 
	    {
	      // Before going insane trying to find the cause of an error, please check if the destination file is not open.	
	      LogHandler.getInstance().writeLog("ERROR: " + ex.getMessage());
	      LogHandler.getInstance().writeLog("Execution aborted.");
	      System.exit(4);
	    } 
	  }
	
}
