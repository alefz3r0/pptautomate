package com.gatta.pptautomate.utils;

import org.apache.poi.sl.usermodel.PictureData.PictureType;

public class Base64Image {
	private PictureType type;
	private byte[] data;
	
	public Base64Image(byte[] data, PictureType type) {
		this.type = type;
		this.data = data;
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
}
