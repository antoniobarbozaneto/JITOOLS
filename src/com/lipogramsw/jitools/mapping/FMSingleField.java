package com.lipogramsw.jitools.mapping;

import java.util.UUID;

import org.apache.commons.codec.binary.Base64;


public class FMSingleField {

	//private int fieldIndex = 0;							// Internal index
	//private int fieldInputIndex = 0;					// Input index (order in input file, starts in ZERO)
	//private int fieldOutputIndex = 0;					// Output index (order in output file, starts in ZERO)
	//private String fieldOutputName = "";				// Output name (presentation)
	
	
	
	private String autoType = FMTypes.AUTO_UUID;		// Automatic value generator type (default is UUID)
	
	
	public FMSingleField() {
		// TODO Auto-generated constructor stub
	}
	
	public String generateAutoValue()
	{
		String autoValue = new String("");
		
		
		if (this.autoType.equalsIgnoreCase(FMTypes.AUTO_UUID)) 
		{
			// Returns a Type4 UUID (pseudo-random based)
			autoValue = UUID.randomUUID().toString();
		}
		
		
		return autoValue;
	}
	
	public String encodeToBase64(String originalValue)
	{
		String valueBase64 = new String("");
		
		valueBase64 = new String(Base64.encodeBase64(originalValue.getBytes()));
		
		return valueBase64;
	}

}
