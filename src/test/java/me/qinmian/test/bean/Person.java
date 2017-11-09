package me.qinmian.test.bean;

import java.util.Date;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

import me.qinmian.annotation.DataStyle;
import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;
import me.qinmian.annotation.ExcelRowCell;
import me.qinmian.annotation.ExportCellStyle;
import me.qinmian.annotation.ExportFontStyle;
import me.qinmian.annotation.HeadStyle;
import me.qinmian.annotation.StaticExcelRow;

@StaticExcelRow(cells={@ExcelRowCell(startRow=0,autoCol=true,value="人类表")})
//@ExportStyle( headStyle=@ExportCellStyle(alignment=CellStyle.ALIGN_CENTER,verticalAlignment=CellStyle.VERTICAL_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.DARK_BLUE.index)))
@HeadStyle(@ExportCellStyle(alignment=CellStyle.ALIGN_CENTER,verticalAlignment=CellStyle.VERTICAL_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.DARK_BLUE.index)))
@DataStyle
@Excel(headRow=1,dataRow=2)
public class Person {

	@ExcelField(headName="性别")
	private String gender;
	
	@ExcelField(headName="类型")
	private String type;
	
	@ExcelField(headName="名字")
	private String name;
	
	@ExcelField(headName="生日",dataFormat="yyyy年m月d日",width=15)
	private Date birthday;
	
	@ExcelField(headName="身高")
	private Integer hight;
	
	@ExcelField(headName="昵称")
	private String nickname;
	
	@ExcelField(headName="体重")
	private Integer weight;
	
	@ExcelField(headName="登陆时间",dataFormat="yyyy年m月d日",width=15)
	private Date loginDate;
	
	@ExcelField(headName="电话")
	private String phone;

	public String getGender() {
		return gender;
	}

	public void setGender(String gender) {
		this.gender = gender;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
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

	public Integer getHight() {
		return hight;
	}

	public void setHight(Integer hight) {
		this.hight = hight;
	}

	public String getNickname() {
		return nickname;
	}

	public void setNickname(String nickname) {
		this.nickname = nickname;
	}

	public Integer getWeight() {
		return weight;
	}

	public void setWeight(Integer weight) {
		this.weight = weight;
	}

	public Date getLoginDate() {
		return loginDate;
	}

	public void setLoginDate(Date loginDate) {
		this.loginDate = loginDate;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}
	
}
