package me.qinmian.util;

import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

public class DateFormatHolder {

	private final static ThreadLocal<Map<String,SimpleDateFormat>> holder = new ThreadLocal<Map<String,SimpleDateFormat>>();
	
	public static void put(String formatStr ,SimpleDateFormat format) {
		if(holder.get() == null) {
			holder.set(new HashMap<String,SimpleDateFormat>());
		}
		holder.get().put(formatStr, format);
	}
	
	public static SimpleDateFormat get(String formatStr) {
		return holder.get() == null ? null : holder.get().get(formatStr);
	}
	
	public static void remove() {
		holder.remove();
	}
}
