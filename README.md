# excel-utils

pojo+注解形式一键导出excel
pojo的目前支持的类型为基本数据类型及其包装类型，pojo可以嵌套pojo，但是目前不支持List和Map，pojo可以继承，导出回将父类的属性一并导出。
    目前支持的功能有：导出excel03、07，可设置样式、分sheet，分次导出，在xls

## pojo的定义如下

```Java
@StaticExcelRow(cells={
				@ExcelRowCell(startRow=0,value="${msg}",autoCol=true,cellStyle=@ExportCellStyle(verticalAlignment=CellStyle.VERTICAL_CENTER,
								alignment=CellStyle.ALIGN_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.BLUE.index))),
				@ExcelRowCell(startRow=1,autoCol=true,value="${status}",cellStyle=@ExportCellStyle(verticalAlignment=CellStyle.VERTICAL_CENTER,alignment=CellStyle.ALIGN_CENTER
								,fontStyle=@ExportFontStyle(color=HSSFColor.RED.index)))
				})
@HeadStyle(@ExportCellStyle(alignment=CellStyle.ALIGN_CENTER,verticalAlignment=CellStyle.VERTICAL_CENTER,fontStyle=@ExportFontStyle(color=HSSFColor.DARK_BLUE.index)))
@DataStyle
@Excel(headRow=2,dataRow=5)
public class User {
	
	@ExcelField(headName="用户名",required=true)
	private String username;
	
	@ExcelField(headName="密码")
	private String password;
	
	@ExcelField(headName="邮箱",width=18)
	@DataStyle(@ExportCellStyle(fontStyle=@ExportFontStyle(fontName="微软雅黑",color=HSSFColor.BLUE.index,underline=Font.U_SINGLE)))
	private String email;
	
	@ExcelField(headName="电话")
	private String phone;
	
	@ExcelField(headName="地址")
	@DataStyle(@ExportCellStyle(fontStyle=@ExportFontStyle(color=HSSFColor.RED.index)))
	private String address;
	
	@ExcelField(headName="昵称")
	private String nickName;
	
	@ExcelField(headName="名字")
	private String name;
	
	@ExcelField(headName="生日",dateFormat="yyyy年MM月dd日 HH点",width=15,dataType=DataType.String)
	private Date birthday;
    
    ...省略get，set....
}
```
```Java
@Excel(headRow=2,dataRow=5,sheetName="用户统计表",sheetSize=65536)
public class UserPlus extends User{
	
	@ExcelField(headName="用户ID",sort=101)
	private int id;

	private String firstName;
	
	@ExcelField(headName="角色")
	private Role role;
    ....省略get，set....
}
角色POJO如下：
public class Role {
	
	@ExcelField(headName="角色名称")
	private String roleName;
	
	@ExcelField(headName="角色描述")
	private String roleDesc;

	@IgnoreField
	private Privilege privilege;

    ....省略get，set...
}
```
****
其中：@Excel代表是一个需要导出导入excel的类，可以设置的属性如下：
****
|属性 |含义|
|----|----|
|headRow|表头所在的行，从0开始，默认为0
|dataRow|数据行开始的行数，从0开始，默认为1
|sheetSize|每个sheet显示最大的行数，包括表头和数据以及静态行
|sheetName|sheet显示的名字，实际名字为该(属性i),i表示第几个sheet
****
@ExportStyle表示导出表格的样式，其中可设置的属性如下，值得注意的是该注解可以写在类上和字段之上，类上代表全局的，字段上代表特有的，若字段为pojo类型，该pojo之中得到属性的样式的全局样式即该字段上的样式
****
|属性|含义|
|----|----|
|headStyle|表头的样式，为注解类型
|dataStyle|数据行的样式，为注解类型
|dataEqHead|boolean类型，是否数据行的样式与表头一致，默认为false
|dataHightInPoint|数据行的的高度，单位为字号
|headHightInPoint|表头行的高度，单位为字号
****
==注意：2017-11-09更新之中将表头样式和数据行样式分离，分为两个注解，注解如下：==
****
@HeadStyle和DataStyle，两个注解的属性是一致
****
|属性|含义|
|--|--|
|value|@ExportCellStyle类型，设置具体的样式
|hightInPoint|高度，单位为字号|
****


@StaticExcelRow表示一些与数据无关的行，比如一些介绍信息，属性如下：
****
|属性|含义|
|----|----|
|cells|注解@ExcelRowCell类型数组，是所有静态行集合
****
@ExcelRowCell注解属性如下：
****
|属性|含义|
|----|----|
|value|该静态行需要显示的数据
|startRow|起始行，从开始
|endRow|结束行
|startCol|起始列，从0开始
|endCol|结束列，从0开始
|rowHightInPoint|行高，单位为字号
|startCol|起始列，从0开始
|autoCol|是否自动设置宽度与数据列数相匹配，默认为false
|isSingle|是否是一个单元格，默认为fasle
|cellStyle|单元格样式，为@ExportCellStyle注解类型
****

@ExportCellStyle是与CellStyle之中可以设置的方法一致的

## 分sheet和批次导出实现

1.  分sheet
*****
```Java
@Excel(headRow=2,dataRow=5,sheetName="用户统计表",sheetSize=65536)
```
只要设置sheetSize，然后导出就会自动分sheet，每个sheet为65536行，包括静态行、表投行和数据行，其中xls最大为65536，xlsx为100w，导出示例如下：
```Java
@Test
public void separateSheet() throws IOException {
	//获取10w条数据
	List<UserPlus> list = getData(0);
	Map<String, String> map = new HashMap<String,String>();
	map.put("msg", "用户信息导出报表");
	map.put("status", "导出成功");
	long start = System.currentTimeMillis();
	Workbook workbook = ExcelExportUtil.exportExcel03(UserPlus.class, list, map);
	long end = System.currentTimeMillis();
	System.out.println("耗时：" + (end - start ) + "毫秒");
	FileOutputStream outputStream = new FileOutputStream("D:/test/user.xls");
	workbook.write(outputStream);
	outputStream.flush();
	outputStream.close();
}
```
分批导出代码示例如下：
```Java
@Test
public void batchesExport() throws IOException {
	Map<String, String> map = new HashMap<String,String>();
	map.put("msg", "用户信息导出报表");
	map.put("status", "导出成功");
	
	Workbook workbook = null;
	//分两次写入20w条数据
	List<UserPlus> list;
	for(int i = 0 ; i < 2 ; i++){
		list = getData(i*100000);
		workbook = ExcelExportUtil.exportExcel03(UserPlus.class, list, map,workbook);
		
	}
	FileOutputStream outputStream = new FileOutputStream("D:/test/user1.xls");
	workbook.write(outputStream);
	outputStream.flush();
	outputStream.close();
}
```

导出表头部分效果图如下：
![image](http://chuantu.biz/t6/126/1509846351x1902307777.png)
分shett：
![image](http://chuantu.biz/t6/126/1509846554x1902307777.png)
20w条数据被分为4个sheet

