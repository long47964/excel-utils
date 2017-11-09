package me.qinmian.test.processor;

import me.qinmian.bean.inter.ImportProcessor;

public class MyImportProcessor implements ImportProcessor{

	@Override
	public Object process(Object val) {
		return "唐家三少";
	}

}
