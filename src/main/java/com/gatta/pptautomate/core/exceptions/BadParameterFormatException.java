package com.gatta.pptautomate.core.exceptions;

@SuppressWarnings("serial")
public class BadParameterFormatException extends Exception { 
	public BadParameterFormatException(String errorMessage) {
		super(errorMessage);
	}
}
