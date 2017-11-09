package me.qinmian.test.processor;

import me.qinmian.bean.inter.LinkProcessor;
import me.qinmian.emun.LinkType;

public class MyLinkProcessor implements LinkProcessor{

	@Override
	public Object process(Object fieldVal, Object currentVal) {
		return fieldVal;
	}

	@Override
	public LinkType getLinkType() {
		return LinkType.Email;
	}

	@Override
	public String getLinkAddress(Object fieldVal, Object currentVal) {
		return fieldVal.toString();
	}

	

}
