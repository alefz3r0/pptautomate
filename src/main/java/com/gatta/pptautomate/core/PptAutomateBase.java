package com.gatta.pptautomate.core;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.SequenceInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.PictureData.PictureType;

import com.gatta.pptautomate.core.exceptions.BadParameterFormatException;
import com.gatta.pptautomate.core.utils.Base64Image;
import com.gatta.pptautomate.core.utils.Position;
import com.gatta.pptautomate.core.utils.Size;

import groovy.json.JsonSlurper;
import groovy.lang.Binding;
import groovy.lang.GroovyShell;

abstract class PptAutomateBase implements PptAutomateInterface {
	
	Logger logger = LogManager.getLogger(PptAutomateBase.class);
	
	private Binding binding;
	private GroovyShell shell;
	
	public PptAutomateBase() {
		binding = new Binding();
		shell = new GroovyShell(binding);
		binding.setVariable("outputPpt", this);
	}
	
	public final PptAutomateBase withAppendTemplateSlides(Integer start, Integer stop) {
		ArrayList<Integer> idx = new ArrayList<>();
		
		//TODO check if start<stop
		for (int i = start; i<= stop; i++) {
			idx.add(i);
		}
		
		return this.withAppendTemplateSlides(idx);
	}
	
	public final PptAutomateBase withAppendTemplateSlide(Integer slide) {
		return withAppendTemplateSlides(slide, slide);
	}
	
	public final PptAutomateBase selectOutputSlides(int from, int to) {
		ArrayList<Integer> tmp = new ArrayList<>();
		for (int i = from; i <= to; i++) {
			tmp.add(i);
		}
		
		return selectOutputSlides(tmp);
	}
	
	protected abstract void checkTargetSlideIdx(List<Integer> idx);
	
	protected abstract void resetTargetShapes();
	
	public final PptAutomateBase selectOutputSlide(Integer i) {
		return selectOutputSlides(i, i);
	}

	public final PptAutomateBase selectAllOutputSlides() {
		//Check no output slides 
		if (getOutputPptSlidesCount() == 0) throw new IllegalStateException("No output slides have been created yet - nothing to select");
		
		return selectOutputSlides(1, getOutputPptSlidesCount());
	}
	
	public final PptAutomateBase selectAllShapes() {
		return selectShapesMatchingRegex(".*");
	}
	
	public final PptAutomateBase selectShapes(String name) {
		return selectShapesMatchingRegex(Pattern.quote(name));
	}
	
	public final PptAutomateBase fillColor(String color) {
		Color c = null;
		if (color != null) {
			try {
				c = PptUtils.getColor(color);
			} catch (BadParameterFormatException e) {
				return this;
			}
		}
		
		return fillColor(c);
	}
	
	public final PptAutomateBase replaceWithImg(Base64Image img) {
		return replaceWithImg(img, true);
	}
	
	public final PptAutomateBase move(Position position) {
		return move(position, false);
	}
	
	public final PptAutomateBase resize(Size size) {
		return resize(size, false);
	}
	
	public final PptAutomateBase processText() {
		return processText(false);
	}
	
	public final Binding getBinding() {
		return binding;
	}
	
	public final GroovyShell getGroovyShell() {
		return shell;
	}
	
	public final PptAutomateBase executeGroovyScript(InputStream scriptIs) {
		//GroovyShell shell = new GroovyShell(getBinding());
    	
    	String importStr = "";
    	importStr += "import " + PptAutomateBase.class.getName() + ";";
    	importStr += "import " + Base64Image.class.getName() + ";";
    	importStr += "import " + Position.class.getName() + ";";
    	importStr += "import " + Size.class.getName() + ";";
    	importStr += "import " + PictureType.class.getName() + ";";
    	
    	Vector<InputStream> streams = new Vector<>();
    	streams.add(new ByteArrayInputStream(importStr.getBytes(StandardCharsets.UTF_8)));
    	streams.add(scriptIs);
    	streams.add(new ByteArrayInputStream("return outputPpt".getBytes(StandardCharsets.UTF_8)));
    	
    	//TODO check cast?
    	return (PptAutomateBase)shell.evaluate(new InputStreamReader(new SequenceInputStream(streams.elements())));
	}
	
	public final void parseJsonToBinding(String variableName, InputStream input) {
		getBinding().setVariable(variableName, new JsonSlurper().parse(input));
	}
}
