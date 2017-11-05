package me.qinmian.bean;

import java.lang.reflect.Method;
import java.util.List;

public class ImportFieldInfo {

	private List<Class<?>> typeChain;
	
	private List<Method> setMethodChain;
	
	private List<Method> getMethodChain;
	
	private boolean required;

	private String dateFormat;
	
	
	public ImportFieldInfo() {}


	public ImportFieldInfo(List<Class<?>> typeChain, List<Method> setMethodChain, List<Method> getMethodChain,
			boolean required, String dateFormat) {
		super();
		this.typeChain = typeChain;
		this.setMethodChain = setMethodChain;
		this.getMethodChain = getMethodChain;
		this.required = required;
		this.dateFormat = dateFormat;
	}


	public List<Class<?>> getTypeChain() {
		return typeChain;
	}


	public void setTypeChain(List<Class<?>> typeChain) {
		this.typeChain = typeChain;
	}


	public List<Method> getSetMethodChain() {
		return setMethodChain;
	}


	public void setSetMethodChain(List<Method> setMethodChain) {
		this.setMethodChain = setMethodChain;
	}


	public List<Method> getGetMethodChain() {
		return getMethodChain;
	}


	public void setGetMethodChain(List<Method> getMethodChain) {
		this.getMethodChain = getMethodChain;
	}


	public boolean isRequired() {
		return required;
	}


	public void setRequired(boolean required) {
		this.required = required;
	}


	public String getDateFormat() {
		return dateFormat;
	}


	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}

	
}
