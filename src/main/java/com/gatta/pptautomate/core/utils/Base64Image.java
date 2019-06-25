package com.gatta.pptautomate.core.utils;

import java.io.UnsupportedEncodingException;
import java.util.Base64;

import org.apache.poi.sl.usermodel.PictureData.PictureType;

public class Base64Image {
	private PictureType type;
	private byte[] data;
	
	public Base64Image(byte[] data, PictureType type) {
		this.type = type;
		this.data = data;
	}
	
	public Base64Image(byte[] data, String type) {
		this.type = parseMimeType(type);
		this.data = data;
	}
	
	public Base64Image(String data, PictureType type) throws UnsupportedEncodingException {
		this.type = type;
		this.data = Base64.getDecoder().decode(data.getBytes("UTF-8"));
	}
	
	public Base64Image(String data, String type) throws UnsupportedEncodingException {
		this.type = parseMimeType(type);
		this.data = Base64.getDecoder().decode(data.getBytes("UTF-8"));
	}
	
	public PictureType getType() {
		return type;
	}
	public void setType(PictureType type) {
		this.type = type;
	}
	public byte[] getData() {
		return data;
	}
	public void setData(byte[] data) {
		this.data = data;
	}
	
	public static PictureType parseMimeType(String mimeType) {
		switch(mimeType) {
			case "image/bmp":
				return PictureType.BMP;
			case "image/gif":
				return PictureType.GIF;
			case "image/jpeg":
				return PictureType.JPEG;
			case "image/png":
				return PictureType.PNG;
			case "image/tiff":
				return PictureType.TIFF;
			default:
				throw new IllegalArgumentException("MimeType not supported: " + mimeType);
		}
	}
}
