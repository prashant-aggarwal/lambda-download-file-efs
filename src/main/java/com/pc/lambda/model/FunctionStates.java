package com.pc.lambda.model;

import lombok.Data;

@Data
public class FunctionStates {
	
	private boolean documentFound;
	private boolean fileDownloaded;
	private boolean dataFormatted;
	private boolean dataValidated;
	private boolean dataMatched;
	private boolean sourceMatched;
	private boolean excelFormatNeeded;
	private boolean excelCalcNeeded;
	private boolean reportPrepared;
}
