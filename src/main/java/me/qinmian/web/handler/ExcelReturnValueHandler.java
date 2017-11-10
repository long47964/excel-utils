package me.qinmian.web.handler;

import java.io.BufferedOutputStream;
import java.io.InputStream;
import java.util.UUID;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.MethodParameter;
import org.springframework.util.StringUtils;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;

import me.qinmian.emun.ExcelFileType;
import me.qinmian.web.annotation.ExcelFile;

public class ExcelReturnValueHandler implements HandlerMethodReturnValueHandler {

	private String sourceCharset = "UTF-8";

	private String targetCharset = "ISO8859-1";
	
	@Override
	public boolean supportsReturnType(MethodParameter returnType) {
		return returnType.hasMethodAnnotation(ExcelFile.class) && 
				(Workbook.class.isAssignableFrom(returnType.getMethod().getReturnType()) || 
				InputStream.class.isAssignableFrom(returnType.getMethod().getReturnType())
				);

	}

	@Override
	public void handleReturnValue(Object returnValue, MethodParameter returnType, ModelAndViewContainer mavContainer,
			NativeWebRequest webRequest) throws Exception {
		BufferedOutputStream bos = null;
		Workbook workbook = null;
		mavContainer.setRequestHandled(true);
		try {
			String suffix = ".xls";
			ExcelFile excelFile = returnType.getMethodAnnotation(ExcelFile.class);
			String exportName = excelFile.value();
			if(StringUtils.isEmpty(exportName)){
				exportName = UUID.randomUUID().toString();
			}
			if(Workbook.class.isAssignableFrom(returnValue.getClass())){
				workbook = (Workbook) returnValue;
			}else{
				InputStream stream = (InputStream) returnValue;
				if(excelFile.excelType() == ExcelFileType.XLS){
					workbook = new HSSFWorkbook(stream);
				}else{
					workbook = new XSSFWorkbook(stream);
					suffix = ".xlsx";
				}
			}
			HttpServletResponse resp = webRequest.getNativeResponse(HttpServletResponse.class);
			resp.reset();
			resp.setContentType("application/vnd.ms-excel;charset=utf-8");
			resp.setHeader("Content-Disposition", "attachment;filename="+ 
							new String((exportName + suffix).getBytes(sourceCharset), targetCharset));
			bos = new BufferedOutputStream(resp.getOutputStream());
			workbook.write(bos);
			bos.flush();
		} catch (Exception e) {
			throw e;
		}finally {
			if(bos != null ){
				bos.close();
			}
			if(workbook != null && SXSSFWorkbook.class.equals(workbook.getClass())){
				SXSSFWorkbook wb = (SXSSFWorkbook) workbook;
				wb.dispose();
			}
		}
	}

	public String getSourceCharset() {
		return sourceCharset;
	}

	public void setSourceCharset(String sourceCharset) {
		this.sourceCharset = sourceCharset;
	}

	public String getTargetCharset() {
		return targetCharset;
	}

	public void setTargetCharset(String targetCharset) {
		this.targetCharset = targetCharset;
	}

}
