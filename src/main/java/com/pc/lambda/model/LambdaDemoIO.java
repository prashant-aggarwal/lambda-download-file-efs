package com.pc.lambda.model;

import lombok.Data;

@Data
public class LambdaDemoIO {
	
	private String firstName;
	private String lastName;
	private FunctionStates funcStates;
	private FileDetails[] fileDetails;
}
