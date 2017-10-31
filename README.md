# excel-utils

pojo+注解形式一键导出excel
pojo的目前支持的类型为基本数据类型及其包装类型，pojo可以嵌套pojo，但是目前不支持List和Map，pojo可以继承，导出回将父类的属性一并导出

##pojo的定义如下

```java(type)
@StaticExcelRow(cells={@ExcelRowCell(startRow=0,autoCol=true,value="人类表")})
@ExportStyle( headStyle=@ExportCellStyle(alignment=CellStyle.ALIGN_CENTER,verticalAlignment=CellStyle.VERTICAL_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.DARK_BLUE.index)))
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
    
    ...省略get，set....
}
```

****
其中：@Excel代表是一个需要导出导入excel的类，可以设置的属性如下：
|属性 | 含义|
|----|----|
|headRow|表头所在的行，从0开始，默认为0|
|dataRow|数据行开始的行数，从0开始，默认为1|
|sheetSize|每个sheet显示最大的行数，包括表头和数据以及静态行|
|sheetName|sheet显示的名字，实际名字为该(属性i),i表示第几个sheet|

@ExportStyle表示导出表格的样式，其中可设置的属性如下，值得注意的是该注解可以写在类上和字段之上，类上代表全局的，字段上代表特有的，若字段为pojo类型，该pojo之中得到属性的样式的全局样式即该字段上的样式
属性|含义
|----|----|
|headStyle|表头的样式，为注解类型|
|dataStyle|数据行的样式，为注解类型|
|dataEqHead|boolean类型，是否数据行的样式与表头一致，默认为false|
|dataHightInPoint|数据行的的高度，单位为字号|
|headHightInPoint|表头行的高度，单位为字号|


