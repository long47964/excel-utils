package me.qinmian.bean;

public class ExportFontStyleInfo {
		
	private Short boldweight;
	
	private Short color;
	
//	private Short fontHeight;
	
	private Short fontHeightInPoints;
	
	private String fontName;
	
	private Boolean italic;
	
	private Boolean strikeout;
	
	private Short typeOffset;
	
	private Byte underline;

	public Short getBoldweight() {
		return boldweight;
	}

	public void setBoldweight(Short boldweight) {
		this.boldweight = boldweight;
	}

	public Short getColor() {
		return color;
	}

	public void setColor(Short color) {
		this.color = color;
	}

	/*public Short getFontHeight() {
		return fontHeight;
	}

	public void setFontHeight(Short fontHeight) {
		this.fontHeight = fontHeight;
	}*/

	public Short getFontHeightInPoints() {
		return fontHeightInPoints;
	}

	public void setFontHeightInPoints(Short fontHeightInPoints) {
		this.fontHeightInPoints = fontHeightInPoints;
	}

	public String getFontName() {
		return fontName;
	}

	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	public Boolean getItalic() {
		return italic;
	}

	public void setItalic(Boolean italic) {
		this.italic = italic;
	}

	public Boolean getStrikeout() {
		return strikeout;
	}

	public void setStrikeout(Boolean strikeout) {
		this.strikeout = strikeout;
	}

	public Short getTypeOffset() {
		return typeOffset;
	}

	public void setTypeOffset(Short typeOffset) {
		this.typeOffset = typeOffset;
	}

	public Byte getUnderline() {
		return underline;
	}

	public void setUnderline(Byte underline) {
		this.underline = underline;
	}
	
}
