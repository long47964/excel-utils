package me.qinmian.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import me.qinmian.test.bean.Person;
import me.qinmian.util.ExcelExportUtil;

public class TestPerson {

	@Test
	public void exportExcel() throws FileNotFoundException, Exception {
		File file = new File("D:/test/test/ppp44.xls");
		FileOutputStream outputStream = new FileOutputStream(file);
		List<Person> list = getData();
		Workbook workbook = null;
		long start = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel03(Person.class, list);
//		workbook = ExcelExportUtil.exportExcel(Person.class,list,ExcelFileType.XLS,null,null,true);
//		list = getData();
//		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLSX,map,workbook,true);
		long end = System.currentTimeMillis();
		
		System.out.println("耗时：" + (end-start) + "毫秒");		
		System.out.println("****************************");
		workbook.write(outputStream);
		
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
