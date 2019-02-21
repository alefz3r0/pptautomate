package com.gatta.pptautomate.exceptions;

@SuppressWarnings("serial")
public class BadParameterFormatException extends Exception { 
	public BadParameterFormatException(String errorMessage) {
		super(errorMessage);
	}
}
