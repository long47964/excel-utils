package me.qinmian.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import me.qinmian.emun.ExcelFileType;
import me.qinmian.test.bean.Role;
import me.qinmian.test.bean.UserPlus;
import me.qinmian.util.ExcelExportUtil;

public class TestUser {

	/**
	 * 分批写出20w条数据
	 * 
	 * @throws IOException 
	 */
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


	/**@throws IOException 
	 * @Excel(headRow=2,dataRow=5,sheetName="用户统计表",sheetSize=65536)
	 * 分sheet测试  
	 */
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
	
	@Test
	public void exportExcel() throws FileNotFoundException, Exception {
		File file = new File("D:/test/test/test4.xls");
		FileOutputStream outputStream = new FileOutputStream(file);
		List<UserPlus> list = getData(0);
		Map<String,String> map = new HashMap<String,String>();
		map.put("msg", "来自星星的你的说得好的");
		map.put("status", "显示成功");
		Workbook workbook = null;
		long start = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLS,map,null,false);
//		list = getData(1798);
//		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list,ExcelFileType.XLS,map,workbook,false);
		long end = System.currentTimeMillis();
		
		System.out.println("耗时：" + (end-start) + "毫秒");		
		workbook.write(outputStream);
		list.clear();
		
		list = getData(544);
		long start1 = System.currentTimeMillis();
		workbook = ExcelExportUtil.exportExcel(UserPlus.class,list, ExcelFileType.XLS,map,null,false);
		long end1 = System.currentTimeMillis();
		System.out.println("耗时：" + (end1-start1) + "毫秒");		
		FileOutputStream out = new FileOutputStream("D:/test/test/test33.xls");
		workbook.write(out);
		
		out.flush();
		out.close();
		
		//SXSSFWorkbook需要手动调用dispose，这样才能删除临时文件
		if(SXSSFWorkbook.class.equals(workbook.getClass())){
			SXSSFWorkbook wb = (SXSSFWorkbook)workbook;
			wb.dispose();
		}
		outputStream.flush();
		outputStream.close();
		
	}

	/** 产生10w条数据
	 * @param start
	 * @return
	 */
	private static List<UserPlus> getData(int start) {
		List<UserPlus> list = new ArrayList<UserPlus>();
		for(int i = start ; i < 100000 + start ; i++){
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
