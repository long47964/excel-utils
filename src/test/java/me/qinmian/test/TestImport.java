package me.qinmian.test;

import java.io.File;
import java.io.FileInputStream;
import java.util.Collection;
import java.util.List;

import org.junit.Test;

import me.qinmian.test.bean.BookShelf;
import me.qinmian.util.ExcelImportUtil;
import me.qinmian.util.ImportResult;

public class TestImport {

	@Test
	public void importBookResult() throws Exception {
		File file = new File("D:/test/book.xls");
		String fileName = file.getName();
		FileInputStream fileInputStream = new FileInputStream(file);
		long start = System.currentTimeMillis();
		ImportResult<BookShelf> result = ExcelImportUtil.importExcel(fileName, fileInputStream, BookShelf.class);
		long end = System.currentTimeMillis();
		System.out.println("耗时：" + (end-start) + "毫秒");
		Collection<BookShelf> collection = result.getDataMap().values();
		System.out.println(collection);
	}
	
	@Test
	public void importBookList() throws Exception {
		File file = new File("D:/test/book.xls");
		String fileName = file.getName();
		FileInputStream fileInputStream = new FileInputStream(file);
		long start = System.currentTimeMillis();
		List<BookShelf> list = ExcelImportUtil.importExcel(BookShelf.class, fileName, fileInputStream);
		long end = System.currentTimeMillis();
		System.out.println("耗时：" + (end-start) + "毫秒");
		System.out.println(list);
	}
}
