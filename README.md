# excel-utils

pojo+注解形式一键导入导出excel

pojo的目前支持的类型为基本数据类型及其包装类型，pojo可以嵌套pojo，但是目前pojo字段不支持List和Map，pojo可以继承，导出会将父类的属性一并导出。
    
导出目前支持的功能有：导出excel03、07，根据pojo定义形成合并表头，导出类型设定，设置样式、分sheet，分批导出。

导入目前只支持单行表头，pojo定义可嵌套。


### pojo的定义如下

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

### 导出使用的类和方法
类名：ExcelExportUtil
****
|方法名|作用|重载参数|
|------|----|--------|
|exportExcel03|导出xls格式的excel表格|clazz,list,staticRowData,workbook|
|exportExcel07|导出xlsx格式的excel表格|比上面多了一个可传参数：xlsxQuickMode|
****
各个参数的意义
****
|名称|类型|意义|
|----|----|----|
|clazz|Class<T>|pojo的Class|
|list|List<T>|数据集合|
|staticRowData|Map<String,String>|当静态行的值为${key}时，将从这个map之中获取数据写出|
|workbook|Workbook|Workbook对象，分批写出时，需要传入|
|xlsxQuickMode|boolean|仅在导出格式为xlsx有用，默认为false
****
注意：xlsxQuickMode为true时使用SXSSFWorkbook导出，在数据量大的时候，速度比XSSFWorkbook快得多，但是过程之中会产生临时文件，临时文件需要调用dispose()方法释放删除。

### 导入使用的类和方法
类名：ExcelImportUtil
|方法名|作用|重载参数|
|------|----|--------|
|importExcel|导入excel|clazz,fileName,inputStream,fileType|
各个参数的意义
****
|名称|类型|意义|
|----|----|----|
|clazz|Class<T>|pojo的Class|
|fileName|String|文件名|
|inputStream|InputStream|输入流|
|fileType|ExcelFileType|枚举类型，表示excel格式|
****
导入返回类型
|类型|属性|
|----|----|
|List<T>|pojo数据集合|
|ImportResult<T>|自定义结果对象，包含成功和失败条数，失败的行号，以及一个行号为key，对应数据的map|

****
导出调用示例
```Java
List<UserPlus> list = getData(0, 100000);
Map<String, String> map = new HashMap<String,String>();
map.put("msg", "用户信息导出报表");
map.put("status", "导出成功");
long start = System.currentTimeMillis();
Workbook workbook = ExcelExportUtil.exportExcel03(UserPlus.class, list, map);
```
导出表格图片如下：
![image](http://chuantu.biz/t6/126/1509846351x1902307777.png)

导入示例
```Java
File file = new File("D:/test/book.xls");
String fileName = file.getName();
FileInputStream fileInputStream = new FileInputStream(file);
long start = System.currentTimeMillis();
ImportResult<BookShelf> result = ExcelImportUtil.importExcel(fileName, fileInputStream, BookShelf.class);
```