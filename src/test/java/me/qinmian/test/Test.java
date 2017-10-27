package me.qinmian.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import me.qinmian.emun.ExcelFileType;
import me.qinmian.test.bean.Role;
import me.qinmian.test.bean.User;
import me.qinmian.test.bean.UserPlus;
import me.qinmian.util.ExcelExportUtil;
import me.qinmian.util.ExcelImportUtil;
import me.qinmian.util.ImportResult;

public class Test {

	public static void main(String[] args) throws Exception {
		exportExcel();
	}

	@SuppressWarnings("unused")
	private static void importExcel() throws FileNotFoundException, Exception {
		File file = new File("D:/test/test3.xlsx");
		String fileName = file.getName();
		FileInputStream fileInputStream = new FileInputStream(file);
		long start = System.currentTimeMillis();
		ImportResult<User> result = ExcelImportUtil.importExcel(fileName, fileInputStream, User.class);
		long end = System.currentTimeMillis();
		System.out.println(result);
		System.out.println("耗时：" + (end-start) + "毫秒");
	}
	
	private static void exportExcel() throws FileNotFoundException, Exception {
		File file = new File("D:/test/test/test4.xls");
		FileOutputStream outputStream = new FileOutputStream(file);
		List<UserPlus> list = getData(0);
		Map<String,String> map = new HashMap<String,String>();
		map.put("msg", "来自星星的你的说得好的");
		map.put("status", "显示成功");
		Workbook workbook = null;
		long start = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLS,map,null,false);
		list = getData(1798);
		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLS,map,workbook,false);
		long end = System.currentTimeMillis();
		
		System.out.println("耗时：" + (end-start) + "毫秒");		
		workbook.write(outputStream);
		/*list.clear();
		
		list = getData(544);
		long start1 = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list, ExcelFileType.XLSX,map,null,false);
		long end1 = System.currentTimeMillis();
		System.out.println("耗时：" + (end1-start1) + "毫秒");		
		FileOutputStream out = new FileOutputStream("D:/test/test/test33.xlsx");
		workbook.write(out);
		
		out.flush();
		out.close();*/
		
		if(SXSSFWorkbook.class.equals(workbook.getClass())){
			SXSSFWorkbook wb = (SXSSFWorkbook)workbook;
			wb.dispose();
		}
		outputStream.flush();
		outputStream.close();
		
	}

	private static List<UserPlus> getData(int start) {
		List<UserPlus> list = new ArrayList<UserPlus>();
		for(int i = start ; i < 50000 + start ; i++){
			UserPlus user = new UserPlus();
			user.setAddress("深圳"+i);
			user.setEmail("111qq"+ i + "@qq.com");
			user.setBirthday(new Date());
			user.setPassword("qinmian" + i);
			user.setUsername("qin"+i);
			user.setPhone("155344"+i);
			user.setId(i);
			user.setFirstName("first");
			user.setNickName("nickname"+i);
			user.setName("name"+i);
			user.setRole(new Role("role"+i,"desc"+i));
			list.add(user);
		}
		return list;
	}

}
