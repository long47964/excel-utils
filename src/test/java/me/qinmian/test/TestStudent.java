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

import me.qinmian.emun.ExcelFileType;
import me.qinmian.test.bean.Role;
import me.qinmian.test.bean.StudentEntity;
import me.qinmian.util.ExcelExportUtil;

public class TestStudent {

	@Test
	public void exportExcel() throws FileNotFoundException, Exception {
		File file = new File("D:/test/test/stu4.xls");
		FileOutputStream outputStream = new FileOutputStream(file);
		List<StudentEntity> list = getData();
		Workbook workbook = null;
		long start = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(StudentEntity.class,list,ExcelFileType.XLS,null,null,false);
//		list = getData();
//		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLSX,map,workbook,true);
		long end = System.currentTimeMillis();
		
		System.out.println("耗时：" + (end-start) + "毫秒");		
		System.out.println("****************************");
		workbook.write(outputStream);
		list.clear();
		
		list = getData();
		long start1 = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(StudentEntity.class,list, ExcelFileType.XLS,null,null,false);
		long end1 = System.currentTimeMillis();
		System.out.println("耗时：" + (end1-start1) + "毫秒");	
		System.out.println("****************************");
		FileOutputStream out = new FileOutputStream("D:/test/test/stu33.xls");
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
	
	private static List<StudentEntity> getData() {
		List<StudentEntity> list = new ArrayList<StudentEntity>();
		StudentEntity s;
		for (int i = 0; i < 50000; i++) {
			s = new StudentEntity();
			s.setId(""+i);
			s.setRegistrationDate(new Date());
			s.setName("name"+i);
			s.setSex("男");
			s.setBirthday(new Date());
			s.setUsername("username"+i);
			s.setPassword("password"+i);
			s.setEmail("email"+i);
			s.setAddress("深圳"+i);
			s.setRole(new Role("name","desc"));
			list.add(s);
		}
		return list;
	}
	
}
