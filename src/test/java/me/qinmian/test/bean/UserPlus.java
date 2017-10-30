package me.qinmian.test.bean;

import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;

@Excel(headRow=2,dataRow=5,sheetName="用户统计表",sheetSize=30000)
public class UserPlus extends User{
	
	@ExcelField(headName="用户ID",sort=101)
	private int id;

	private String firstName;
	
	@ExcelField(headName="角色")
	private Role role;
	
	public Role getRole() {
		return role;
	}

	public void setRole(Role role) {
		this.role = role;
	}

	public UserPlus() {
		
	}
	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	
	
}
