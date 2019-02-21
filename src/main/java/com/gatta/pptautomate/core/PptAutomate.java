package com.gatta.pptautomate.core;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.Reader;
import java.io.SequenceInputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Vector;
import java.util.stream.Collectors;

import javax.swing.text.html.parser.ParserDelegator;

import org.apache.commons.io.IOUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.sl.usermodel.TextShape.TextAutofit;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSimpleShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlString;
import org.codehaus.groovy.runtime.GStringImpl;

import com.gatta.pptautomate.core.exceptions.BadParameterFormatException;
import com.gatta.pptautomate.core.utils.Base64Image;
import com.gatta.pptautomate.core.utils.Position;
import com.gatta.pptautomate.core.utils.PptHtmlParser;
import com.gatta.pptautomate.core.utils.PptUtils;
import com.gatta.pptautomate.core.utils.Size;

import groovy.lang.Binding;
import groovy.lang.GroovyShell;

public class PptAutomate {

	private XMLSlideShow output = null;
	private Integer templateSlidesCount;
	private List<Integer> targetSlides = new ArrayList<>();
	private List<XSLFShape> targetShapes = new ArrayList<>();
	
	private Binding binding = new Binding();

	Logger logger = LogManager.getLogger(PptAutomate.class);

	public PptAutomate(InputStream templateIS) throws IOException {
		logger.debug("Instantiating PptAutomate object");
		InputStream is;
		byte[] templateBytes = null;

		try {
			templateBytes = IOUtils.toByteArray(templateIS);

			is = new ByteArrayInputStream(templateBytes);
			this.output = new XMLSlideShow(is);
			is.close();
		} catch (IOException e) {
			logger.error("Cannot read the template");
			throw new IOException(e.getMessage());
		}

		this.templateSlidesCount = output.getSlides().size();
		
    	binding.setVariable("outputPpt", this);
		
		logger.debug("PptAutomate object instantiated");
	}
	
	/* COPY SLIDES FROM TEMPLATE METHODS */
	public PptAutomate withAppendTemplateSlides(ArrayList<Integer> templateSlidesIdx) {
		//Check idx out of bound
		List<Integer> oob = templateSlidesIdx.stream().filter(x -> x >= templateSlidesCount+1).collect(Collectors.toList());
		if (oob.size() > 0) {
			throw new IndexOutOfBoundsException("Template indices out of bound: " + oob.toString());
		}
		
		//Check idx length
		if (templateSlidesIdx.size() == 0) {
			throw new IllegalArgumentException("No indexes provided");
		}
		
		//Copy target slides at the end
		for (int i : templateSlidesIdx) {
			try {
				logger.debug("Copying template slide #{} to output slide #{}", i, output.getSlides().size()+1-templateSlidesCount);
				XSLFSlideLayout targetSlideLayout = output.getSlides().get(i-1).getSlideLayout();
				XSLFSlide targetSlide = output.getSlides().get(i-1);
				output.createSlide(targetSlideLayout).importContent(targetSlide);
				//TODO manage exceptions in order to verify the template
			} catch (IndexOutOfBoundsException e) {
				logger.error("Template index out of bound: {}", i);
				throw new IndexOutOfBoundsException(e.getMessage());
			}
		}
		
		selectOutputSlides(output.getSlides().size()-templateSlidesCount-templateSlidesIdx.size()+1, output.getSlides().size()-templateSlidesCount);
				
		return this;
	}
	
	/* OUTPUT SLIDES SELECT METHODS */
	public PptAutomate selectOutputSlides(int from, int to) {
		List<Integer> tmp = new ArrayList<>();
		for (int i = from; i <= to; i++) {
			tmp.add(i);
		}
		
		checkTargetSlideIdx(tmp);
		
		logger.debug("Output slides selected from #{} to #{}", from, to);
		targetSlides = tmp;
		resetTargetShapes();
		
		return this;
	}
	
	public PptAutomate selectOutputSlides(ArrayList<Integer> slidesIdx) {
		checkTargetSlideIdx(slidesIdx);
		
		logger.debug("Output slides selected: {}", slidesIdx);
		targetSlides = slidesIdx;
		resetTargetShapes();
		
		return this;
	}
	
	public PptAutomate selectOutputSlide (Integer i) {
		return selectOutputSlides(i, i);
	}
	
	public PptAutomate selectAllOutputSlides() {
		//Check no output slides
		int lastOutputSlide = output.getSlides().size() - templateSlidesCount; 
		if (lastOutputSlide == 0) throw new IllegalStateException("No output slides have been created yet - nothing to select");
		
		return selectOutputSlides(1, lastOutputSlide);
	}
	
	private void checkTargetSlideIdx(List<Integer> idx) {
		//Check out of bounds
		List<Integer> oob = idx.stream().filter(x -> x > output.getSlides().size() - templateSlidesCount).collect(Collectors.toList());
		if (oob.size() > 0) throw new IndexOutOfBoundsException("Tried to select output slides with index out of bounds: " + oob.toString());
		
		//Check idx length
		if (idx.size() == 0) {
			throw new IllegalArgumentException("No indexes provided");
		}
	}
	
	private int slideNumToIdx(int slideNum) {
		return slideNum + templateSlidesCount - 1;
	}
	
	/* SHAPES SELECT METHODS */
	public PptAutomate selectShapes(String name) {
		logger.debug("Selecting shapes within target slides with name: {}", name);
		
		resetTargetShapes();
		
		for (int i : targetSlides) {
			output.getSlides().get(slideNumToIdx(i)).getShapes().stream()
				.filter(x -> x.getShapeName().equals(name))
				.forEach(targetShapes::add);			
		}
		
		logger.debug("Found: {} shapes", targetShapes.size());
		
		return this;
	}
	
	public PptAutomate selectShapesMatchingRegex(String regex) {
		logger.debug("Selecting shapes within target slides with name matching regex: {}", regex);
		
		resetTargetShapes();
		
		for (int i : targetSlides) {
			output.getSlides().get(slideNumToIdx(i)).getShapes().stream()
				.filter(x -> x.getShapeName().matches(regex))
				.forEach(targetShapes::add);			
		}
		
		logger.debug("Found: {} shapes", targetShapes.size());
		
		return this;
	}
	
	public PptAutomate selectAllShapes() {
		logger.debug("Selecting all shapes within target slides");
		
		resetTargetShapes();
		
		for (int i : targetSlides) {
			output.getSlides().get(slideNumToIdx(i)).getShapes().stream()
				.forEach(targetShapes::add);			
		}
		
		logger.debug("Found: {} shapes", targetShapes.size());
		
		return this;
	}
	
	private void resetTargetShapes() {
		targetShapes = new ArrayList<>();
	}
	
	/* SHAPE ACTIONS METHODS */
	
	public PptAutomate fillColor(String color) {
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
	
	public PptAutomate fillColor(Color color) {
		logger.debug("Filling shapes with color: {}", color);
		
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFSimpleShape) {
				((XSLFSimpleShape) s).setFillColor(color);
			} else logger.warn("Shape {} is not istance of XSLFSimpleShape: cannot process action Fill", s.getShapeName());
		}
		
		return this;
	}
	
	public PptAutomate replaceWithImg(Base64Image img, Boolean keepAspectRatio) {
		//TODO testare bene il workaround e commentare, poi ottimizzare il codice
		
		logger.debug("Replacing shapes with image");
		
		for (XSLFShape s : targetShapes) {
			Rectangle2D rect = s.getAnchor();
			XSLFPictureData p = getPptPictureData(img);
	
			Integer zOrder = getShapeZOrder(s);
	
			//First create new shape, then clone slide
			XSLFSlide originalSlide = (XSLFSlide) s.getSheet();
			XSLFPictureShape s2 = originalSlide.createPicture(p);
			
			XSLFSlide cloneSlide = output.createSlide().importContent(originalSlide);
	
			//Remove shapes from original
			for (int i = originalSlide.getShapes().size()-2; i >= zOrder; i--) {
				originalSlide.removeShape(originalSlide.getShapes().get(zOrder));
			}
	
			//Remove shapes from clone
			cloneSlide.removeShape(cloneSlide.getShapes().get(cloneSlide.getShapes().size()-1));
			for (int i = zOrder; i >= 0; i--) {
				cloneSlide.removeShape(cloneSlide.getShapes().get(i));
			}
	
			//Merge slides and remove clone
			originalSlide.appendContent(cloneSlide);
			output.removeSlide(output.getSlides().size()-1);
	
			if (!keepAspectRatio) {
				s2.setAnchor(rect);
			} else {
				Rectangle2D r = s2.getAnchor();
				if (rect.getWidth() < rect.getHeight()) {
					r.setRect(
							rect.getX(),
							rect.getY(),
							rect.getWidth(),
							s2.getAnchor().getHeight()*rect.getWidth()/s2.getAnchor().getWidth());
				} else {
					r.setRect(
							rect.getX(),
							rect.getY(),
							s2.getAnchor().getWidth()*rect.getHeight()/s2.getAnchor().getHeight(),
							rect.getHeight());
	
				}
				s2.setAnchor(r);
			}
		}
		return this;
	}
	
	public PptAutomate replaceWithImg(Base64Image img) {
		return replaceWithImg(img, true);
	}
	
	private XSLFPictureData getPptPictureData(byte[] pd, PictureType pt) {
		for (XSLFPictureData p : output.getPictureData()) {
			if (Arrays.equals(p.getData(), pd)) {
				//img exists, no need to add
				return p;
			}
		}
		XSLFPictureData p = output.addPicture(pd, pt);

		return p;
	}

	@SuppressWarnings("unused")
	private XSLFPictureData getPptPictureData(Base64Image img) {
		return getPptPictureData(img.getData(), img.getType());
	}

	@SuppressWarnings("unused")
	private XSLFPictureData getPptPictureData(XSLFPictureData img) {
		return getPptPictureData(img.getData(), img.getType());
	}

	private Integer getShapeZOrder(XSLFShape s) {
		Integer out = 0;
		for (XSLFShape s2 : s.getParent().getShapes()) {
			if (s2.getShapeId() == s.getShapeId()) return out;
			out++;
		}
		return null;
	}
	
	public PptAutomate move(Position position, Boolean relative) {
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFSimpleShape) {
				Rectangle2D r = s.getAnchor();
				if (relative) {
					r.setRect(r.getX() + position.getX(), r.getY() + position.getY(), r.getWidth(), r.getHeight());
				} else {
					r.setRect(position.getX(), position.getY(), r.getWidth(), r.getHeight());
				}
				((XSLFSimpleShape)s).setAnchor(r);
			} else logger.warn("Shape {} is not istance of XSLFSimpleShape: cannot process action Move", s.getShapeName());
		}
		return this;
	}
	
	public PptAutomate move(Position position) {
		return move(position, false);
	}
	
	public PptAutomate resize(Size size, Boolean relative) {
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFSimpleShape) {
				Rectangle2D r = s.getAnchor();
				if (relative) {
					r.setRect(r.getX(), r.getY(), r.getWidth() + size.getW(), r.getHeight() + size.getH());
				} else {
					r.setRect(r.getX(), r.getY(), size.getW(), size.getH());
				}
				((XSLFSimpleShape)s).setAnchor(r);
				//TODO manage other units other than pixels, also for MOVE
			} else logger.warn("Shape {} is not istance of XSLFSimpleShape: cannot process action Resize", s.getShapeName());
		}
		return this;
	}

	public PptAutomate resize(Size size) {
		return resize(size, false);
	}
	
	public PptAutomate setTextHtml(String string) {
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFTextShape) {
				Reader r = new StringReader(string);
				
				ParserDelegator pd = new ParserDelegator();			
				try {
					pd.parse(r, new PptHtmlParser((XSLFTextShape)s), false);
				} catch (IOException e) {
					logger.error("Cannot parse HTML string: ");
				}
	
				try {
					r.close();
				} catch (IOException e) {
					logger.warn("Cannot close the string reader");
				}
	
				//Autofit fix
				if (((XSLFTextShape) s).getTextAutofit() == TextAutofit.NORMAL) {
					//Autofit text
					Rectangle2D rect = s.getAnchor();
					((XSLFTextShape) s).resizeToFitText();
	
					while (s.getAnchor().getHeight() > rect.getHeight()) {
						((XSLFTextShape) s).resizeToFitText();
						for (XSLFTextParagraph p : ((XSLFTextShape) s).getTextParagraphs()) {
							for (XSLFTextRun run : p.getTextRuns()) {
								run.setFontSize(run.getFontSize()-.5);
							}
						}
					}
					((XSLFTextShape) s).setAnchor(rect);
				} else if (((XSLFTextShape) s).getTextAutofit() == TextAutofit.SHAPE) {
					//Autofit shape
					((XSLFTextShape) s).resizeToFitText();
				}
			} else logger.warn("Shape {} is not istance of XSLFTextShape: cannot process action SetTextHtml", s.getShapeName());
		}
		return this;
	}

	public PptAutomate processText() {
		GroovyShell g = new GroovyShell(binding);
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFTextShape) {
				for (XSLFTextParagraph p : ((XSLFTextShape) s).getTextParagraphs()) {
					for (XSLFTextRun r : p.getTextRuns()) {
						r.setText(
							((GStringImpl)g.evaluate('"' + r.getRawText().replaceAll("\"","\\\"") + '"' )).toString()
						);
					}
				}
				
				//Workaround for modifying XML
				XmlObject[] xmlTexts = s.getXmlObject().selectPath("declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:t");
				for (XmlObject t : xmlTexts)
				if (t instanceof XmlString) {
					((XmlString) t).setStringValue(((GStringImpl)g.evaluate('"' + ((XmlString) t).getStringValue().replaceAll("\"","\\\"") + '"' )).toString());
				}
	
				//Autofit fix
				if (((XSLFTextShape) s).getTextAutofit() == TextAutofit.NORMAL) {
					//Autofit text
					Rectangle2D rect = s.getAnchor();
					((XSLFTextShape) s).resizeToFitText();
	
					while (s.getAnchor().getHeight() > rect.getHeight()) {
						((XSLFTextShape) s).resizeToFitText();
						for (XSLFTextParagraph p : ((XSLFTextShape) s).getTextParagraphs()) {
							for (XSLFTextRun run : p.getTextRuns()) {
								run.setFontSize(run.getFontSize()-.5);
							}
						}
					}
					((XSLFTextShape) s).setAnchor(rect);
				} else if (((XSLFTextShape) s).getTextAutofit() == TextAutofit.SHAPE) {
					//Autofit shape
					((XSLFTextShape) s).resizeToFitText();
				}
			} else logger.warn("Shape {} is not istance of XSLFTextShape: cannot process action ProcessText", s.getShapeName());
		}
		return this;		
	}
	
	public PptAutomate delete() {
		for (XSLFShape s : targetShapes) {
			s.getSheet().removeShape(s);
		}
	
		return this;
	}
	
	/* FINALIZE PPT METHODS */
	
	public void finalizeAndWritePpt(OutputStream os) throws IOException {
		//Remove template slides
		for (int i = 0; i<templateSlidesCount; i++) {
			output.removeSlide(0);
		}

		try {
			logger.info("Writing output ppt");
			output.write(os);
			output.close();
		} catch (IOException e) {
			logger.error("Cannot write the output ppt to the provided OutputStream");
			throw new IOException(e.getMessage());
		}
	}
	
	/* OTHER METHODS */
	public List<Integer> getTargetSlides() {
		return targetSlides;
	}
	
	public List<XSLFShape> getTargetShapes() {
		return targetShapes;
	}
	
	public Binding getBinding() {
		return binding;
	}
	
	public PptAutomate executeGroovyScript(InputStream scriptIs) {
		GroovyShell shell = new GroovyShell(binding);
    	
    	String importStr = "";
    	importStr += "import " + PptAutomate.class.getName() + ";";
    	
    	Vector<InputStream> streams = new Vector<>();
    	streams.add(new ByteArrayInputStream(importStr.getBytes(StandardCharsets.UTF_8)));
    	streams.add(scriptIs);
    	streams.add(new ByteArrayInputStream("return outputPpt".getBytes(StandardCharsets.UTF_8)));
    	
    	//TODO check cast?
    	return (PptAutomate)shell.evaluate(new InputStreamReader(new SequenceInputStream(streams.elements())));
	}

}
