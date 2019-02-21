package com.gatta.pptautomate.core.utils;

import java.awt.Color;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.gatta.pptautomate.core.exceptions.BadParameterFormatException;

public class PptUtils {

	static Logger logger = LogManager.getLogger(PptUtils.class);
	
	public static Color getColor(String s) throws BadParameterFormatException {
		Color out = null;
		try {
			out = (Color)Color.class.getField(s).get(null);
		} catch (IllegalArgumentException | IllegalAccessException | NoSuchFieldException
				| SecurityException e) {
			try {
				//Try to parse RGB instead
				Pattern p= Pattern.compile("rgb *\\( *([0-9]+), *([0-9]+), *([0-9]+) *\\)");
				Matcher m = p.matcher(s);

				if (m.matches()) {
					out = new Color(Integer.valueOf(m.group(1)),  	// r
							Integer.valueOf(m.group(2)),  			// g
							Integer.valueOf(m.group(3))); 			// b 
				} else {
					//Try hex format
					out = new Color(
							Integer.valueOf( s.substring( 1, 3 ), 16 ),
							Integer.valueOf( s.substring( 3, 5 ), 16 ),
							Integer.valueOf( s.substring( 5, 7 ), 16 ) );
				}
			} catch (Exception e2) {
				logger.error("Color parsing error for string: {}", s);
				throw new BadParameterFormatException("Color parsing error for string: " + s);
			}
		}
		return out;
	}
}
