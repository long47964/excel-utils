package me.qinmian.bean;

public class SortableName {
	
	private Integer sort;
	
	private String name;
	
	public SortableName() {}

	public SortableName(Integer sort, String name) {
		this.sort = sort;
		this.name = name;
	}

	public Integer getSort() {
		return sort;
	}

	public void setSort(Integer sort) {
		this.sort = sort;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	
	
}
