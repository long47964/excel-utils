package me.qinmian.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import me.qinmian.emun.ExcelFileType;
import me.qinmian.test.bean.Person;
import me.qinmian.util.ExcelExportUtil;

public class TestPerson {

	public static void main(String[] args) throws Exception {

		exportExcel();
	}

	private static void exportExcel() throws FileNotFoundException, Exception {
		File file = new File("D:/test/test/ppp44.xls");
		FileOutputStream outputStream = new FileOutputStream(file);
		List<Person> list = getData();
		Workbook workbook = null;
		long start = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(Person.class,list,ExcelFileType.XLS,null,null,true);
//		list = getData();
//		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLSX,map,workbook,true);
		long end = System.currentTimeMillis();
		
		System.out.println("耗时：" + (end-start) + "毫秒");		
		System.out.println("****************************");
		workbook.write(outputStream);
		list.clear();
		
		list = getData();
		long start1 = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(Person.class,list, ExcelFileType.XLS,null,null,true);
		long end1 = System.currentTimeMillis();
		System.out.println("耗时：" + (end1-start1) + "毫秒");	
		System.out.println("****************************");
		FileOutputStream out = new FileOutputStream("D:/test/test/pp33.xls");
		workbook.write(out);
		
		out.flush();
		out.close();
		
		if(SXSSFWorkbook.class.equals(workbook.getClass())){
			SXSSFWorkbook wb = (SXSSFWorkbook)workbook;
			wb.dispose();
		}
		outputStream.flush();
		outputStream.close();
		
	}
	
	private static List<Person> getData() {
		List<Person> list = new ArrayList<Person>();
		Person p;
		for (int i = 0; i < 50000; i++) {
			p = new Person();
			p.setBirthday(new Date());
			p.setGender("男");
			p.setHight(i);
			p.setWeight(i);
			p.setLoginDate(new Date());
			p.setName("name" + i);
			p.setNickname("nick" + i);
			p.setPhone("183"+i);
			p.setType("type" + i);
			
			list.add(p);
		}
		return list;
	}
	
}
