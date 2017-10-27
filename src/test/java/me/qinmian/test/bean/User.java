package me.qinmian.test.bean;

import java.util.Date;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;
import me.qinmian.annotation.ExcelRowCell;
import me.qinmian.annotation.ExportCellStyle;
import me.qinmian.annotation.ExportFontStyle;
import me.qinmian.annotation.ExportStyle;
import me.qinmian.annotation.StaticExcelRow;


@StaticExcelRow(cells={
				@ExcelRowCell(startRow=0,value="${msg}",autoCol=true,cellStyle=@ExportCellStyle(verticalAlignment=CellStyle.VERTICAL_CENTER,
								alignment=CellStyle.ALIGN_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.BLUE.index))),
				@ExcelRowCell(startRow=1,autoCol=true,value="${status}",cellStyle=@ExportCellStyle(verticalAlignment=CellStyle.VERTICAL_CENTER,alignment=CellStyle.ALIGN_CENTER
								,fontStyle=@ExportFontStyle(color=HSSFColor.RED.index)))
				})
@ExportStyle( dataHightInPoint=25 ,headStyle=@ExportCellStyle(alignment=CellStyle.ALIGN_CENTER,verticalAlignment=CellStyle.VERTICAL_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.DARK_BLUE.index)))
@Excel(headRow=2,dataRow=5)
public class User {
	
	@ExcelField(headName="用户名",required=true)
	private String username;
	
	@ExcelField(headName="密码")
	private String password;
	
	@ExcelField(headName="邮箱",width=18)
	private String email;
	
	@ExcelField(headName="电话")
	private String phone;
	
	@ExcelField(headName="地址")
	@ExportStyle(dataStyle=	@ExportCellStyle(
			fontStyle=@ExportFontStyle(fontName="微软雅黑",color=HSSFColor.RED.index,underline=Font.U_SINGLE)))
	private String address;
	
	@ExcelField(headName="昵称")
	private String nickName;
	
	@ExcelField(headName="名字")
	private String name;
	
	@ExcelField(headName="生日",dataFormat="yyyy年m月d日",width=15)
	private Date birthday;

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

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	
	public String getNickName() {
		return nickName;
	}

	public void setNickName(String nickName) {
		this.nickName = nickName;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	@Override
	public String toString() {
		return "User [username=" + username + ", password=" + password + ", email=" + email + ", phone=" + phone
				+ ", address=" + address + ", birthday=" + birthday + "]";
	}
	
	
}
