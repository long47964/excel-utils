package me.qinmian.bean;

import java.util.Comparator;

import me.qinmian.bean.inter.Sortable;

public class SortComparator implements Comparator<Sortable>{

	private static SortComparator comparator;
	
	public int compare(Sortable o1, Sortable o2) {		
		return o1.getSort() - o2.getSort();
	}
	
	public static SortComparator getInstance(){
		return comparator == null ? new SortComparator() : comparator;
	}
}
