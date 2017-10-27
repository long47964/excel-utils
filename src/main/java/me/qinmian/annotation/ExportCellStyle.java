package me.qinmian.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

/**  
 * 各个值参考CellStyle类的静态值ֵ
 * @see org.apache.poi.ss.usermodel.CellStyle
 * 颜色值参看HSSFColor中的颜色值
 * @see org.apache.poi.hssf.util.HSSFColor
 * @author qinmian
 *
 */

@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExportCellStyle {
		
	short alignment() default CellStyle.ALIGN_CENTER;//
	
	short verticalAlignment() default CellStyle.VERTICAL_CENTER;
	
	short borderBottom() default CellStyle.BORDER_NONE;
	
	short borderLeft() default CellStyle.BORDER_NONE;
	
	short borderRight() default CellStyle.BORDER_NONE;
	
	short borderTop() default CellStyle.BORDER_NONE;
	
	short bottomBorderColor() default HSSFColor.BLACK.index;
	
	short leftBorderColor() default HSSFColor.BLACK.index;
	
	short rightBorderColor() default HSSFColor.BLACK.index;
	
	short topBorderColor() default HSSFColor.BLACK.index;
	
	short fillBackgroundColor() default HSSFColor.WHITE.index;
	
	short fillForegroundColor() default HSSFColor.WHITE.index;
	
	short fillPattern() default CellStyle.NO_FILL;
	
	/** 缩进字符数
	 * @return
	 */
	short indention() default 0;
	
	/** 可写值-90到90 ,回转
	 * @return
	 */
	short rotation() default 0 ;
	
	boolean hidden() default false;
	
	boolean locked() default false;
	
	/** 
	 * @return
	 */
	boolean wrapText() default false;
		
	ExportFontStyle fontStyle() default @ExportFontStyle;
	
}
