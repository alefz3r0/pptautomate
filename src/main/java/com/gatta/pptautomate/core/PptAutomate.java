package com.gatta.pptautomate.core;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.gatta.pptautomate.core.utils.Base64Image;
import com.gatta.pptautomate.core.utils.Position;
import com.gatta.pptautomate.core.utils.Size;

import groovy.lang.Binding;
import groovy.lang.GroovyShell;

public class PptAutomate implements PptAutomateInterface {

	PptAutomateBase pptAutomate;
	
	Logger logger = LogManager.getLogger(PptAutomate.class);
	
	public enum PptLibrary {
		APACHE_POI
	}
	
	public PptAutomate(InputStream templateIS) throws IOException {
		//Defaults to Apache POI
		this(templateIS, PptLibrary.APACHE_POI);
	}
	
	public PptAutomate(InputStream templateIS, PptLibrary pptLibrary) throws IOException {
		switch (pptLibrary) {
		case APACHE_POI:
			logger.debug("Using Library: Apache POI");
			this.pptAutomate = new PptAutomateApachePOI(templateIS);
			break;
		}
	}

	@Override
	public PptAutomateBase withAppendTemplateSlides(ArrayList<Integer> templateSlidesIdx) {
		return pptAutomate.withAppendTemplateSlides(templateSlidesIdx);
	}

	@Override
	public PptAutomateBase selectShapes(String name) {
		return pptAutomate.selectShapes(name);
	}

	@Override
	public PptAutomateBase selectShapesMatchingRegex(String regex) {
		return pptAutomate.selectShapesMatchingRegex(regex);
	}

	@Override
	public PptAutomateBase selectAllShapes() {
		return pptAutomate.selectAllShapes();
	}

	@Override
	public PptAutomateBase fillColor(Color color) {
		return pptAutomate.fillColor(color);
	}

	@Override
	public PptAutomateBase replaceWithImg(Base64Image img, Boolean keepAspectRatio) {
		return pptAutomate.replaceWithImg(img, keepAspectRatio);
	}

	@Override
	public PptAutomateBase move(Position position, Boolean relative) {
		return pptAutomate.move(position, relative);
	}

	@Override
	public PptAutomateBase resize(Size size, Boolean relative) {
		return pptAutomate.resize(size, relative);
	}

	@Override
	public PptAutomateBase setTextHtml(String string) {
		return pptAutomate.setTextHtml(string);
	}

	@Override
	public PptAutomateBase processText(Boolean processHtml) {
		return pptAutomate.processText(processHtml);
	}

	@Override
	public PptAutomateBase delete() {
		return pptAutomate.delete();
	}

	@Override
	public void finalizeAndWritePpt(OutputStream os) throws IOException {
		pptAutomate.finalizeAndWritePpt(os);
	}

	@Override
	public List<String> getTargetShapes() {
		return pptAutomate.getTargetShapes();
	}

	@Override
	public PptAutomateBase logTextShapeProperties() {
		return pptAutomate.logTextShapeProperties();
	}

	@Override
	public Integer getOutputPptSlidesCount() {
		return pptAutomate.getOutputPptSlidesCount();
	}

	@Override
	public PptAutomateBase selectOutputSlides(ArrayList<Integer> slidesIdx) {
		return pptAutomate.selectOutputSlides(slidesIdx);
	}

	@Override
	public List<Integer> getTargetSlides() {
		return pptAutomate.getTargetSlides();
	}
	
	@Override
	public Binding getBinding() {
		return pptAutomate.getBinding();
	}
	
	@Override
	public GroovyShell getGroovyShell() {
		return pptAutomate.getGroovyShell();
	}
	
	@Override
	public void parseJsonToBinding(String variableName, InputStream input) {
		pptAutomate.parseJsonToBinding(variableName, input);
	}
	
	@Override
	public PptAutomateBase executeGroovyScript(InputStream scriptIs) {
		return pptAutomate.executeGroovyScript(scriptIs);
	}

	@Override
	public PptAutomateBase selectAllOutputSlides() {
		return pptAutomate.selectAllOutputSlides();
	}
}
