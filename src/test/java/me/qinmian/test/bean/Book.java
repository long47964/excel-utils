package me.qinmian.test.bean;

import java.util.Date;

import me.qinmian.annotation.ExcelField;

public class Book {

	@ExcelField(headName="书名")
	private String name;
	
	@ExcelField(headName="作者")
	private String author;
	
	@ExcelField(headName="发布日期")
	private Date publish;
	
	@ExcelField(headName="价格")
	private Double price;
	
	@ExcelField(headName="数量")
	private Integer total;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public Date getPublish() {
		return publish;
	}

	public void setPublish(Date publish) {
		this.publish = publish;
	}

	public Double getPrice() {
		return price;
	}

	public void setPrice(Double price) {
		this.price = price;
	}

	public Integer getTotal() {
		return total;
	}

	public void setTotal(Integer total) {
		this.total = total;
	}

	@Override
	public String toString() {
		return "Book [name=" + name + ", author=" + author + ", publish=" + publish + ", price=" + price + ", total="
				+ total + "]";
	}
	
}
