package me.qinmian.bean;

import java.util.Comparator;

public class NameComparator implements Comparator<SortableName>{

	public int compare(SortableName o1, SortableName o2) {
		return o1.getSort() - o2.getSort();
	}


}
