package com.lipogramsw.jitools;

import java.text.SimpleDateFormat;

public class LogHandler {


	private static LogHandler logHandler = null;
	private boolean useTimestamp = false; 
	private String logID = "";
	
	
	public static LogHandler getInstance()
	{
		if (logHandler == null) logHandler = new LogHandler();
		return logHandler; 
	}
	
	public void setLogID(String logId)
	{
		this.logID = logId;
	}
	
	public void setTimeStamp(boolean useTimestamp)
	{
		this.useTimestamp = useTimestamp;
	}
	
	public void writeLog(String logLine) 
	  {
	    String logPrefix = "";
	    
	    if (this.useTimestamp) 
	    {
	      logPrefix = String.valueOf(logPrefix) + (new SimpleDateFormat("yyyy-MM-dd HH:mm:ss")).format(Long.valueOf(System.currentTimeMillis())) + " ";
	    }
	    if (!this.logID.trim().isEmpty()) 
	    {
	      logPrefix = String.valueOf(logPrefix) + "ID " + this.logID.trim() + " ";
	    }
	    System.out.println(String.valueOf(logPrefix.trim()) + ": " + logLine);
	  }
	
	
}
