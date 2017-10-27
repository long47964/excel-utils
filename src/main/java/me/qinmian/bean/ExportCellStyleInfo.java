package me.qinmian.bean;


public class ExportCellStyleInfo {

	private Short alignment;//水平显示设置
	
	private Short verticalAlignment;//垂直显示设置
	
	private Short borderBottom; //
	
	private Short borderLeft;//
	
	private Short borderRight;//
	
	private Short borderTop;//
	
	private Short bottomBorderColor;//ֵ
	
	private Short leftBorderColor;
	
	private Short rightBorderColor;
	
	private Short topBorderColor;
	
	private Short fillBackgroundColor;
	
	private Short fillForegroundColor;
	
	private Short fillPattern;
	
	/** 
	 * @return
	 */
	private Short indention;
	
	/** 
	 * @return
	 */
	private Short rotation;
	
	private Boolean hidden;
	
	private Boolean locked;
	
	/** 
	 * @return
	 */
	private Boolean wrapText;
	
	private ExportFontStyleInfo fontStyleInfo;

	public Short getAlignment() {
		return alignment;
	}

	public void setAlignment(Short alignment) {
		this.alignment = alignment;
	}

	public Short getVerticalAlignment() {
		return verticalAlignment;
	}

	public void setVerticalAlignment(Short verticalAlignment) {
		this.verticalAlignment = verticalAlignment;
	}

	public Short getBorderBottom() {
		return borderBottom;
	}

	public void setBorderBottom(Short borderBottom) {
		this.borderBottom = borderBottom;
	}

	public Short getBorderLeft() {
		return borderLeft;
	}

	public void setBorderLeft(Short borderLeft) {
		this.borderLeft = borderLeft;
	}

	public Short getBorderRight() {
		return borderRight;
	}

	public void setBorderRight(Short borderRight) {
		this.borderRight = borderRight;
	}

	public Short getBorderTop() {
		return borderTop;
	}

	public void setBorderTop(Short borderTop) {
		this.borderTop = borderTop;
	}

	public Short getBottomBorderColor() {
		return bottomBorderColor;
	}

	public void setBottomBorderColor(Short bottomBorderColor) {
		this.bottomBorderColor = bottomBorderColor;
	}

	public Short getLeftBorderColor() {
		return leftBorderColor;
	}

	public void setLeftBorderColor(Short leftBorderColor) {
		this.leftBorderColor = leftBorderColor;
	}

	public Short getRightBorderColor() {
		return rightBorderColor;
	}

	public void setRightBorderColor(Short rightBorderColor) {
		this.rightBorderColor = rightBorderColor;
	}

	public Short getTopBorderColor() {
		return topBorderColor;
	}

	public void setTopBorderColor(Short topBorderColor) {
		this.topBorderColor = topBorderColor;
	}

	public Short getFillBackgroundColor() {
		return fillBackgroundColor;
	}

	public void setFillBackgroundColor(Short fillBackgroundColor) {
		this.fillBackgroundColor = fillBackgroundColor;
	}

	public Short getFillForegroundColor() {
		return fillForegroundColor;
	}

	public void setFillForegroundColor(Short fillForegroundColor) {
		this.fillForegroundColor = fillForegroundColor;
	}

	public Short getFillPattern() {
		return fillPattern;
	}

	public void setFillPattern(Short fillPattern) {
		this.fillPattern = fillPattern;
	}

	public Short getIndention() {
		return indention;
	}

	public void setIndention(Short indention) {
		this.indention = indention;
	}

	public Short getRotation() {
		return rotation;
	}

	public void setRotation(Short rotation) {
		this.rotation = rotation;
	}

	public Boolean getHidden() {
		return hidden;
	}

	public void setHidden(Boolean hidden) {
		this.hidden = hidden;
	}

	public Boolean getLocked() {
		return locked;
	}

	public void setLocked(Boolean locked) {
		this.locked = locked;
	}

	public Boolean getWrapText() {
		return wrapText;
	}

	public void setWrapText(Boolean wrapText) {
		this.wrapText = wrapText;
	}

	public ExportFontStyleInfo getFontStyleInfo() {
		return fontStyleInfo;
	}

	public void setFontStyleInfo(ExportFontStyleInfo fontStyleInfo) {
		this.fontStyleInfo = fontStyleInfo;
	}
	
}
