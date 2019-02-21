package com.gatta.pptautomate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.IOUtils;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.config.Configurator;
import org.apache.poi.sl.usermodel.PictureData.PictureType;

import com.gatta.pptautomate.utils.Base64Image;

public class App 
{
    public static void main( String[] args ) throws IOException
    {   
    	@SuppressWarnings("unused")
		Logger logger = LogManager.getLogger(App.class);
    	Configurator.setRootLevel(Level.DEBUG);
    	ClassLoader classloader = Thread.currentThread().getContextClassLoader();
    	
    	PptAutomate pptAutomate = new PptAutomate(classloader.getResourceAsStream("template2.pptx"));
    	
    	pptAutomate.getBinding().setVariable("img", new Base64Image(IOUtils.toByteArray(classloader.getResourceAsStream("test_img.jpg")), PictureType.JPEG));
    	pptAutomate.getBinding().setVariable("productCode", "123");
    	pptAutomate.getBinding().setVariable("num", "4");
    	
    	pptAutomate.executeGroovyScript(classloader.getResourceAsStream("test.groovy"));
    	
    	File file = new File("C:\\Users\\f.gatta\\eclipse-workspace\\pptautomate\\outputs\\output.pptx");
		if (!file.exists()) file.createNewFile();
    	pptAutomate.finalizeAndWritePpt(new FileOutputStream(file));
    }
}
