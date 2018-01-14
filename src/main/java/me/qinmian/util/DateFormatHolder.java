package me.qinmian.util;

import java.text.SimpleDateFormat;

public class DateFormatHolder {

	private static ThreadLocal<SimpleDateFormat> holder = new ThreadLocal<SimpleDateFormat>();
	
	public static void put(SimpleDateFormat format) {
		holder.set(format);
	}
	
	public static SimpleDateFormat get() {
		return holder.get();
	}
	
	public static void remove() {
		holder.remove();
	}
}
