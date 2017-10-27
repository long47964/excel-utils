package me.qinmian.test.bean;

import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;

import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;
import me.qinmian.annotation.ExcelRowCell;
import me.qinmian.annotation.ExportCellStyle;
import me.qinmian.annotation.ExportStyle;
import me.qinmian.annotation.StaticExcelRow;

@StaticExcelRow(cells={@ExcelRowCell(startRow=0,startCol=0,endRow=0,endCol=8,value="计算机一班学生"
		,cellStyle=@ExportCellStyle(verticalAlignment=CellStyle.VERTICAL_CENTER,alignment=CellStyle.ALIGN_CENTER))})
@Excel(sheetName="学生表",dataRow=2,headRow=1)
@ExportStyle(headStyle=@ExportCellStyle(verticalAlignment=CellStyle.VERTICAL_CENTER,alignment=CellStyle.ALIGN_CENTER))
public class StudentEntity {
	
		@ExcelField(headName="ID")
		private String  id;
		
		@ExcelField(headName="名字")
	    private String  name;
		
		@ExcelField(headName="性别")
	    private String  sex;
		
		@ExcelField(headName="生日",dataFormat="yyyy年m月d日",width=15)
	    private Date  birthday;
		
		@ExcelField(headName="注册日期",dataFormat="yyyy年m月d日",width=15)
	    private Date registrationDate;
		
		@ExcelField(headName="用户名")
	    private String  username;
		
		@ExcelField(headName="密码")
	    private String   password;
		
		@ExcelField(headName="邮箱")
	    private String  email;
		
		@ExcelField(headName="地址")
	    private String address;
		
		
		private Role role;

		public String getId() {
			return id;
		}

		public void setId(String id) {
			this.id = id;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public Date getBirthday() {
			return birthday;
		}

		public void setBirthday(Date birthday) {
			this.birthday = birthday;
		}

		public Date getRegistrationDate() {
			return registrationDate;
		}

		public void setRegistrationDate(Date registrationDate) {
			this.registrationDate = registrationDate;
		}

		public String getUsername() {
			return username;
		}

		public void setUsername(String username) {
			this.username = username;
		}

		public String getPassword() {
			return password;
		}

		public void setPassword(String password) {
			this.password = password;
		}

		public String getEmail() {
			return email;
		}

		public void setEmail(String email) {
			this.email = email;
		}

		public String getAddress() {
			return address;
		}

		public void setAddress(String address) {
			this.address = address;
		}

		public String getSex() {
			return sex;
		}

		public void setSex(String sex) {
			this.sex = sex;
		}

		public Role getRole() {
			return role;
		}

		public void setRole(Role role) {
			this.role = role;
		}
	    
	    
}
