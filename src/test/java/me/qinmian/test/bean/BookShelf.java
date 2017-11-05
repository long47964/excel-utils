package me.qinmian.test.bean;

import me.qinmian.annotation.ExcelField;

public class BookShelf {

	@ExcelField(headName="书架品牌")
	private String brand;
	
	@ExcelField(headName="书架高度")
	private Double hight;
	
	@ExcelField(headName="供应号码")
	private String phone;
	
	private Book book;

	public String getBrand() {
		return brand;
	}

	public void setBrand(String brand) {
		this.brand = brand;
	}

	public Double getHight() {
		return hight;
	}

	public void setHight(Double hight) {
		this.hight = hight;
	}

	public Book getBook() {
		return book;
	}

	public void setBook(Book book) {
		this.book = book;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}
	
	
}
