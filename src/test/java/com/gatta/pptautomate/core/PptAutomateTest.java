package com.gatta.pptautomate.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.PipedInputStream;
import java.io.PipedOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSimpleShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com.gatta.pptautomate.core.PptAutomate;

public class PptAutomateTest {

	private PptAutomate outputPpt;
	private static InputStream templateIs;
		
	@AfterAll
	public static void finish() throws IOException {
		templateIs.close();
	}
	
	@BeforeEach
	public void reset() throws IOException {
		templateIs = Thread.currentThread().getContextClassLoader().getResourceAsStream("test_template.pptx");
		outputPpt = new PptAutomate(templateIs);
	}
	
	@Test
	public void copyTemplateSlides() throws IOException {
		ArrayList<Integer> templateSlidesIdx = new ArrayList<>();
		templateSlidesIdx.add(1);
		templateSlidesIdx.add(2);
		
		outputPpt
			.withAppendTemplateSlides(templateSlidesIdx)
			.withAppendTemplateSlides(templateSlidesIdx);
		
		XMLSlideShow outPpt = getOutputPpt(outputPpt);
		
		assertEquals(4, outPpt.getSlides().size());
	}
	
	@Test
	public void copyTemplateSlidesOutOfBoundThrowsException() {
		ArrayList<Integer> templateSlidesIdx = new ArrayList<>();
		templateSlidesIdx.add(3);
		
		assertThrows(IndexOutOfBoundsException.class, ()->{outputPpt.withAppendTemplateSlides(templateSlidesIdx);});
	}
	
	@Test
	public void shapeActionsAppliedOnlyToCopiedTemplateSlides() throws IOException {
		ArrayList<Integer> templateSlidesIdx = new ArrayList<>();
		templateSlidesIdx.add(1);
		templateSlidesIdx.add(2);
		
		outputPpt
			.withAppendTemplateSlides(templateSlidesIdx)
			.withAppendTemplateSlides(templateSlidesIdx)
				.selectShapes("BOX")
					.fillColor("rgb(0,0,0)");
		
		ArrayList<Integer> targetSlides = new ArrayList<>();
		targetSlides.add(3);
		targetSlides.add(4);
		
		XMLSlideShow out = getOutputPpt(outputPpt);
		Color c = new Color(0, 0, 0);
		
		assertEquals(targetSlides, outputPpt.getTargetSlides());
		assertNotEquals(((XSLFSimpleShape)getShapes(out.getSlides().get(0),"BOX").get(0)).getFillColor(), c);
		assertEquals(((XSLFSimpleShape)getShapes(out.getSlides().get(2),"BOX").get(0)).getFillColor(), c);
	}
	
	@Test
	public void cannotAddEmptyTemplateSlides() {
		ArrayList<Integer> templateSlidesIdx = new ArrayList<>();
		
		assertThrows(IllegalArgumentException.class, ()->{
			outputPpt.withAppendTemplateSlides(templateSlidesIdx);			
		});
		
	}
	
	@Test
	public void cannotSelectWithNoOuputSlides() {
		assertThrows(IllegalStateException.class, ()->{
			outputPpt.selectAllOutputSlides();
		});
	}
	
	/* HELPER METHODS */
	
	private XMLSlideShow getOutputPpt(PptAutomate pptAutomate) throws IOException {
		PipedInputStream is = new PipedInputStream();
		PipedOutputStream os = new PipedOutputStream(is);
		
		new Thread(
			new Runnable(){
				public void run(){
					try {
						pptAutomate.finalizeAndWritePpt(os);
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		).start();
		
		XMLSlideShow out = new XMLSlideShow(is);
		
		out.close();
		is.close();
		
		return out;
	}
	
	private List<XSLFShape> getShapes(XSLFSlide s, String name) {
		return s.getShapes().stream().filter(x -> x.getShapeName().equals(name)).collect(Collectors.toList());
	}
}