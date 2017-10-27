package me.qinmian.bean;

import java.lang.reflect.Field;
import java.util.List;

import me.qinmian.bean.inter.Sortable;

public class SortableField implements Sortable{
	
	private int sort;
	
	private Field field;
	
	private List<SortableField> childFileds;
	
	public SortableField() {}
	public SortableField(int sort, Field field, List<SortableField> childFileds) {
		super();
		this.sort = sort;
		this.field = field;
		this.childFileds = childFileds;
	}



	public int getSort() {
		return sort;
	}

	public void setSort(int sort) {
		this.sort = sort;
	}

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

	public List<SortableField> getChildFileds() {
		return childFileds;
	}

	public void setChildFileds(List<SortableField> childFileds) {
		this.childFileds = childFileds;
	}
	
}
