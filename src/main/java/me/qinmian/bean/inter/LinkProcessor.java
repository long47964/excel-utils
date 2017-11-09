package me.qinmian.bean.inter;

import me.qinmian.emun.LinkType;

public interface LinkProcessor extends ExportProcessor{

	/** 链接类型
	 * @see me.qinmian.emun.LinkType
	 * @return
	 */
	LinkType getLinkType();
	
	/** 链接地址，根据类型而定
	 * @param fieldVal 当前字段值
	 * @param currentVal  当前对象
	 * @return
	 */
	String getLinkAddress(Object fieldVal , Object currentVal);
}
