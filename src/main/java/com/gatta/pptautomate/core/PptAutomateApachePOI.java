package com.gatta.pptautomate.core;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
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
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSimpleShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlString;

import com.gatta.pptautomate.core.utils.Base64Image;
import com.gatta.pptautomate.core.utils.Position;
import com.gatta.pptautomate.core.utils.Size;

import groovy.lang.GroovyShell;

class PptAutomateApachePOI extends PptAutomateBase {

	private XMLSlideShow output = null;
	private Integer templateSlidesCount;
	private List<Integer> targetSlides = new ArrayList<>();
	private List<XSLFShape> targetShapes = new ArrayList<>();

	Logger logger = LogManager.getLogger(PptAutomateApachePOI.class);

	public PptAutomateApachePOI(InputStream templateIS) throws IOException {
		logger.debug("Instantiating PptAutomateApachePOI object");
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
		
		logger.debug("PptAutomateApachePOI object instantiated");
	}
	
	/* COPY SLIDES FROM TEMPLATE METHODS */

	@Override
	public PptAutomateApachePOI withAppendTemplateSlides(ArrayList<Integer> templateSlidesIdx) {
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

	@Override
	protected void checkTargetSlideIdx(List<Integer> idx) {
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
	
	public final PptAutomateBase selectOutputSlides(ArrayList<Integer> slidesIdx) {
		checkTargetSlideIdx(slidesIdx);
		
		logger.debug("Output slides selected: {}", slidesIdx);
		targetSlides = slidesIdx;
		resetTargetShapes();
		
		return this;
	}
	
	/* SHAPES SELECT METHODS */

	@Override
	public PptAutomateApachePOI selectShapesMatchingRegex(String regex) {
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
	
	protected void resetTargetShapes() {
		targetShapes = new ArrayList<>();
	}
	
	/* SHAPE ACTIONS METHODS */

	@Override
	public PptAutomateApachePOI fillColor(Color color) {
		logger.debug("Filling shapes with color: {}", color);
		
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFSimpleShape) {
				((XSLFSimpleShape) s).setFillColor(color);
			} else logger.warn("Shape {} is not istance of XSLFSimpleShape: cannot process action Fill", s.getShapeName());
		}
		
		return this;
	}

	@Override
	public PptAutomateApachePOI replaceWithImg(Base64Image img, Boolean keepAspectRatio) {
		//TODO testare bene il workaround e commentare, poi ottimizzare il codice
		
		logger.debug("Replacing shapes with image");
		
		for (XSLFShape s : targetShapes) {
			
			if (s instanceof XSLFPictureShape) {
				Rectangle2D rect = s.getAnchor();
				//XSLFPictureData p = getPptPictureData(img);
				
				try {
					((XSLFPictureShape) s).getPictureData().setData(img.getData());
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				/*
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
				*/
				
				if (!keepAspectRatio) {
					((XSLFPictureShape) s).setAnchor(rect);
				} else {
					Rectangle2D r = s.getAnchor();
					if (rect.getWidth() < rect.getHeight()) {
						r.setRect(
								rect.getX(),
								rect.getY(),
								rect.getWidth(),
								s.getAnchor().getHeight()*rect.getWidth()/s.getAnchor().getWidth());
					} else {
						r.setRect(
								rect.getX(),
								rect.getY(),
								s.getAnchor().getWidth()*rect.getHeight()/s.getAnchor().getHeight(),
								rect.getHeight());
		
					}
					((XSLFPictureShape) s).setAnchor(r);
				}
			}
		}
		return this;
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

	@Override
	public PptAutomateApachePOI move(Position position, Boolean relative) {
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

	@Override
	public PptAutomateApachePOI resize(Size size, Boolean relative) {
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

	@Override
	public PptAutomateApachePOI setTextHtml(String string) {
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFTextShape) {
				setTextHtml(string, (XSLFTextShape)s);
			} else logger.warn("Shape {} is not istance of XSLFTextShape: cannot process action SetTextHtml", s.getShapeName());
		}
		return this;
	}
	
	private void setTextHtml(String string, XSLFTextShape s) {
		Reader r = new StringReader(string);
		
		ParserDelegator pd = new ParserDelegator();			
		try {
			pd.parse(r, new PptHtmlParser(s), false);
		} catch (IOException e) {
			logger.error("Cannot parse HTML string: ");
		}

		try {
			r.close();
		} catch (IOException e) {
			logger.warn("Cannot close the string reader");
		}

		autofitFix(s);
	}

	@Override
	public PptAutomateApachePOI processText(Boolean processHtml) {
		//GroovyShell g = new GroovyShell(getBinding());
		GroovyShell g = getGroovyShell();
		
		for (XSLFShape s : targetShapes) {
			logger.debug("Processing shape: {}", s.getShapeName());
			if (s instanceof XSLFTextShape) {
				if (processHtml) {
					setTextHtml(
							processGString(g, ((XSLFTextShape)s).getText()),
							(XSLFTextShape)s
					);
				} else {
					for (XSLFTextParagraph p : ((XSLFTextShape) s).getTextParagraphs()) {
						for (XSLFTextRun r : p.getTextRuns()) {
							r.setText(processGString(g, ((XSLFTextShape)s).getText()));
						}
					}
					
					//Workaround for modifying XML
					XmlObject[] xmlTexts = s.getXmlObject().selectPath("declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:t");
					for (XmlObject t : xmlTexts)
					if (t instanceof XmlString) {
						((XmlString) t).setStringValue(processGString(g, ((XSLFTextShape)s).getText()));
					}
					
					autofitFix((XSLFTextShape)s);
				}
			} else logger.warn("Shape {} is not istance of XSLFTextShape: cannot process action ProcessText", s.getShapeName());
		}
		return this;		
	}
	
	private String processGString(GroovyShell g, String s) {
		return (g.evaluate('"' + s.replaceAll("\"","\\\"") + '"')).toString();
	}
	
	private void autofitFix(XSLFTextShape s) {
		//Autofit fix
		if (s.getText().equals("")) return;
		if (s.getTextAutofit() == TextAutofit.NORMAL) {
			//Autofit text
			Rectangle2D rect = s.getAnchor();
			s.resizeToFitText();

			while (s.getAnchor().getHeight() > rect.getHeight()) {
				s.resizeToFitText();
				for (XSLFTextParagraph p : s.getTextParagraphs()) {
					for (XSLFTextRun run : p.getTextRuns()) {
						run.setFontSize(run.getFontSize()-.5);
					}
				}
			}
			s.setAnchor(rect);
		} else if (s.getTextAutofit() == TextAutofit.SHAPE) {
			//Autofit shape
			s.resizeToFitText();
		}
	}

	@Override
	public PptAutomateApachePOI delete() {
		for (XSLFShape s : targetShapes) {
			s.getSheet().removeShape(s);
		}
	
		return this;
	}
	
	/* FINALIZE PPT METHODS */

	@Override
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

	@Override
	public List<String> getTargetShapes() {
		return targetShapes.stream().map(x -> x.getShapeName()).collect(Collectors.toList());
	}

	@Override
	public PptAutomateApachePOI logTextShapeProperties() {
		logger.info("Logging Text Shape Properties for {} selected shapes", targetShapes.size());
		for (XSLFShape s : targetShapes) {
			if (s instanceof XSLFTextShape) {
				logger.info("Sheet #{} - Shape {} - {} paragraph(s)", getSheetIndex(s.getSheet()), s.getShapeName(), ((XSLFTextShape) s).getTextParagraphs().size());
				int pIdx = 1;
				for (XSLFTextParagraph p : ((XSLFTextShape) s).getTextParagraphs()) {
					logger.info("\tParagraph #{}: Text: \"{}\", {} Text Run(s)", pIdx, p.getText(), p.getTextRuns().size());
					int rIdx = 1;
					for (XSLFTextRun r : p.getTextRuns()) {
						logger.info("\t\tText Run #{}: Text: \"{}\", Font Family: {}, Font Color: {}, Font Size: {}, Bold: {}, Italic: {}, Underlined: {}",
							rIdx,
							r.getRawText(),
							r.getFontFamily(),
							r.getFontColor().toString(),
							r.getFontSize(),
							r.isBold(),
							r.isItalic(),
							r.isUnderlined()
						);
						rIdx++;
					}
					pIdx++; 
				}
			} else {
				logger.info("Sheet #{} - Shape {} - Not a Text Shape", getSheetIndex(s.getSheet()), s.getShapeName());
			}
		}
		
		return this;
	}
	
	private Integer getSheetIndex(XSLFSheet sheet) {
		int i = 1;
		for (XSLFSheet s : sheet.getSlideShow().getSlides()) {
			if (s.equals(sheet)) return i - templateSlidesCount;
			i++;
		}
		
		return null;
	}

	@Override
	public Integer getOutputPptSlidesCount() {
		return output.getSlides().size() - templateSlidesCount;
	}
	
	@Override
	public List<Integer> getTargetSlides() {
		return targetSlides;
	}
}
