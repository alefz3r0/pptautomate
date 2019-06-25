package com.gatta.pptautomate.core;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import com.gatta.pptautomate.core.utils.Base64Image;
import com.gatta.pptautomate.core.utils.Position;
import com.gatta.pptautomate.core.utils.Size;

import groovy.lang.Binding;
import groovy.lang.GroovyShell;

interface PptAutomateInterface {

	public PptAutomateBase withAppendTemplateSlides(ArrayList<Integer> templateSlidesIdx);
	
	/* OUTPUT SLIDES SELECT METHODS */
	public PptAutomateBase selectOutputSlides(ArrayList<Integer> slidesIdx);
	public PptAutomateBase selectAllOutputSlides();
	
	/* SHAPES SELECT METHODS */
	public PptAutomateBase selectShapes(String name);
	public PptAutomateBase selectShapesMatchingRegex(String regex);
	public PptAutomateBase selectAllShapes();
	
	/* SHAPE ACTIONS METHODS */
	public PptAutomateBase fillColor(Color color);
	public PptAutomateBase replaceWithImg(Base64Image img, Boolean keepAspectRatio);
	public PptAutomateBase move(Position position, Boolean relative);
	public PptAutomateBase resize(Size size, Boolean relative);
	public PptAutomateBase setTextHtml(String string);
	public PptAutomateBase processText(Boolean processHtml);
	public PptAutomateBase delete();
	
	/* FINALIZE PPT METHODS */
	public void finalizeAndWritePpt(OutputStream os) throws IOException;
	
	/* OTHER METHODS */
	public List<String> getTargetShapes();
	public PptAutomateBase logTextShapeProperties();
	public Integer getOutputPptSlidesCount();
	public List<Integer> getTargetSlides();
	public Binding getBinding();
	public GroovyShell getGroovyShell();
	public void parseJsonToBinding(String variableName, InputStream input);
	public PptAutomateBase executeGroovyScript(InputStream scriptIs);
}
