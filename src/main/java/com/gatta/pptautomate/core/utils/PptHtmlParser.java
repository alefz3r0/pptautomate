package com.gatta.pptautomate.core.utils;

import javax.swing.text.MutableAttributeSet;
import javax.swing.text.html.HTML;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.AutoNumberingScheme;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/*
 * MANAGES:
 * 
 * <strong>
 * <em>
 * <u>
 * <ul> (even nested)
 * <ol> (even nested, by default arabicPeriod)
 * <br/>
 * <p>
 * 
 */
public class PptHtmlParser extends HTMLEditorKit.ParserCallback {

	private XSLFTextShape s;
	private XSLFTextParagraph p;
	private XSLFTextRun r;

	private String fontFamily;
	private Double fontSize;
	private PaintStyle fontColor;
	private TextAlign textAlign;
	
	private Boolean isBold = false;
	private Boolean isEm = false;
	private Boolean isU = false;
	private Integer indentLevel = 0;
	
	private BulletType bulletType = BulletType.NONE;
	
	enum BulletType {
		NONE,
		UL,
		OL
	}

	Logger logger = LogManager.getLogger(PptHtmlParser.class);
	
	public PptHtmlParser(XSLFTextShape s) {
		this.s = s;
		XSLFTextParagraph pOld = s.getTextParagraphs().get(0);
		XSLFTextRun rOld = s.getTextParagraphs().get(0).getTextRuns().get(0);
		//TODO manage exception;
		
		fontFamily = rOld.getFontFamily();
		fontSize = rOld.getFontSize();
		fontColor = rOld.getFontColor();
		
		textAlign = pOld.getTextAlign();
		
		s.clearText();
		this.p = s.addNewTextParagraph();
		p.setTextAlign(textAlign);
		this.r = null;
	}

	@Override
	public void handleStartTag(HTML.Tag t, MutableAttributeSet a, int pos) {
		if (t.equals(HTML.Tag.STRONG)) {
			this.isBold = true;
		} else if (t.equals(HTML.Tag.EM)) {
			this.isEm = true;
		} else if (t.equals(HTML.Tag.U)) {
			this.isU = true;
		} else if (t.equals(HTML.Tag.UL)) {
			this.bulletType = BulletType.UL;
			this.indentLevel++;
			switchParagraph();
		} else if (t.equals(HTML.Tag.OL)) {
			this.bulletType = BulletType.OL;
			this.indentLevel++;
			switchParagraph();
		} else if (t.equals(HTML.Tag.LI)) {

		} else if (t.equals(HTML.Tag.P)) {
			switchParagraph();
		} else if (t.equals(HTML.Tag.BODY) || t.equals(HTML.Tag.HTML) || t.equals(HTML.Tag.HEAD)) {
		} else {
			logger.warn("Unsupported start tag {} found while parsing HTML: ignoring", t);
		}
	}

	@Override
	public void handleEndTag(HTML.Tag t, int pos) {
		if (t.equals(HTML.Tag.STRONG)) {
			this.isBold = false;
		} else if (t.equals(HTML.Tag.EM)) {
			this.isEm = false;
		} else if (t.equals(HTML.Tag.U)) {
			this.isU = false;
		} else if (t.equals(HTML.Tag.UL)) {
			this.bulletType = BulletType.NONE;
			this.indentLevel--;
			switchParagraph();
		} else if (t.equals(HTML.Tag.OL)) {
			this.bulletType = BulletType.NONE;
			this.indentLevel--;
			switchParagraph();
		} else if (t.equals(HTML.Tag.LI)) {
			switchParagraph();
		} else if (t.equals(HTML.Tag.P)) {
			switchParagraph();
		} else if (t.equals(HTML.Tag.BODY) || t.equals(HTML.Tag.HTML) || t.equals(HTML.Tag.HEAD)) {
		} else {
			logger.warn("Unsupported end tag {} found while parsing HTML: ignoring", t);
		}
	}

	@Override
	public void handleSimpleTag(HTML.Tag t, MutableAttributeSet a, int pos) {
		if (t.equals(HTML.Tag.BR)) {
			r = p.addLineBreak();
		} else {
			logger.warn("Unsupported simple tag {} found while parsing HTML: ignoring", t);
		}
	}

	@Override
	public void handleText(char[] data, int pos) {
		if (runIsInvalid()) {
			switchRun();
		}
		
		if (r.getRawText() != null) {
			r.setText(r.getRawText() + String.valueOf(data));
		} else {
			r.setText(String.valueOf(data));
		}
	}

	private boolean runIsInvalid() {
		if (r == null ||
				r.getRawText().equals("\n") ||
				r.isBold() != isBold ||
				r.isItalic() != isEm ||
				r.isUnderlined() != isU) return true;
		else return false;
	}

	private void switchRun() {
		this.r = p.addNewTextRun();
		r.setBold(isBold);
		r.setItalic(isEm);
		r.setUnderlined(isU);
		
		r.setFontFamily(fontFamily);
		r.setFontSize(fontSize);
		r.setFontColor(fontColor);
	}
	
	private void switchParagraph() {
		if (p.getTextRuns().size() != 0) {
			this.p = s.addNewTextParagraph();
			p.setTextAlign(textAlign);
			r = null;
		}
		switch (this.bulletType) {
		case UL:
			p.setBullet(true);
			p.setIndentLevel(indentLevel);
			break;
		case OL:
			p.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1);
			p.setIndentLevel(indentLevel);
		case NONE:
			break;
		default:
			break;
		}
	}
}