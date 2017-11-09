package me.qinmian.bean.inter;

public interface ExportProcessor {

	/**
	 * @param fieldVal 当前字段值
	 * @param currentVal 当前对象
	 * @return
	 */
	Object process(Object fieldVal , Object currentVal);
}
