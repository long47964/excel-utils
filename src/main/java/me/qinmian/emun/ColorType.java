package me.qinmian.emun;

import org.apache.poi.hssf.util.HSSFColor;

public enum ColorType {

	RED(HSSFColor.RED.index),
	GREEN(HSSFColor.YELLOW.index),
	BLUE(HSSFColor.BLUE.index);
	
	private short colorVal;

	public short getValue() {
		return colorVal;
	}

	private ColorType(short colorVal) {
		this.colorVal = colorVal;
	}
	
	
}
