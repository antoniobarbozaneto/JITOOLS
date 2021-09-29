package com.lipogramsw.jitools;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;

import com.lipogramsw.jitools.xlsfiles.XLSConverter;


public class JITools
{
  static Options options;
  static CommandLine cmd = null;
  
  public static void main(String[] args) throws Exception 
  {
    DefaultParser defaultParser = new DefaultParser();
    String outputFile = "";
    String outputType = "xls";			// Add 30.04.2021
    
    initOptions();

    try 
    {
      cmd = defaultParser.parse(options, args);
    }
    catch (Exception e) 
    {
      showHelp(e, options);
      System.exit(4);
    } 
    
    if (cmd.hasOption("log-id"))
    {
    	LogHandler.getInstance().setLogID(cmd.getOptionValue("log-id"));
    }

    if (cmd.hasOption("log-timestamp"))
    {
    	LogHandler.getInstance().setTimeStamp(true);
    }
    else
    {
    	LogHandler.getInstance().setTimeStamp(false);
    }
    
    if (cmd.hasOption("xlsx"))	// Add 30.04.2021
    {
    	 outputType = "xlsx";
    } 
    else
    {
    	outputType = "xls";
    }
    
    
    if (cmd.hasOption("output")) 
    {
    	outputFile = cmd.getOptionValue("output");
    }
    else 
    {
    	//outputFile = cmd.getOptionValue("input").replaceAll(".csv", ".xls");				// <- Rem 30.04.2021
    	outputFile = cmd.getOptionValue("input").replaceAll(".csv", "." + outputType);		// <- Add 30.04.2021
    } 
    
    if (cmd.hasOption("use-template") || cmd.hasOption("start-line")) 
    {
      XLSConverter xlsConverter = new XLSConverter();
      xlsConverter.convertFileTemplate(cmd.getOptionValue("input"), outputFile, cmd.getOptionValue("use-template"), Integer.parseInt(cmd.getOptionValue("start-line")), cmd.getOptionValue("sheet-name"), cmd.getOptionValue("separator"), outputType);
    }
    else 
    {
      XLSConverter xlsConverter = new XLSConverter();
      xlsConverter.convertFile(cmd.getOptionValue("input"), outputFile, cmd.getOptionValue("sheet-name"), cmd.getOptionValue("separator"), outputType);
    } 
    
  }
  
  private static void initOptions() 
  {
    options = new Options();
    options.addRequiredOption(null, "input", true, "Input file ");
    options.addOption(null, "xlsx", false, "Generates XLSX instead of XLS");					 	// <- Add 30.04.2021
    options.addOption(null, "output", true, "Output file (uses the input if none provided)");
    options.addOption(null, "separator", true, "Field separator (optional, default is ',')");
    options.addOption(null, "log-id", true, "General use string for log identification");
    options.addOption(null, "log-timestamp", false, "use timestamp in log messages");
    options.addOption(null, "sheet-name", true, "String to use as XLS file sheet name");
    options.addOption(null, "use-template", true, "Use another XLS file as a template");
    options.addOption(null, "start-line", true, "Line to start the data output in XLS");
    options.addOption(null, "sheet-column", true, "Column that indicates the sheet to be used"); 	// <- Add 12.02.2020
  }

  private static void showHelp(Exception e, Options o) 
  {
    HelpFormatter formatter = new HelpFormatter();
    System.out.println("Java Integration Tools for Thomson Reuters\n");
    System.out.println(e.getMessage());
    formatter.printHelp("csvtools", "Converts CSV files to MS Excel (TM) .XLS ", o, "http://www.lipogramsw.com", true);
  }
 
  
}
