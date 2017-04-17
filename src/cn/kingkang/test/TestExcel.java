package cn.kingkang.test;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import cn.kingkang.office.until.ExcelUntil;
import cn.kingkang.pojo.ExcelSheet;

public class TestExcel {
	private static File file;
	@Before
	public void setUp() throws Exception {
		
	}

	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void test() {
		fail("Not yet implemented");
	}

	@Test
	public void testExcelRead(){
		file = new File("C:\\点名册.xls");
		try {
			List<ExcelSheet> excelSheets = ExcelUntil.getExcelData(file, null, 1);
			System.out.println(excelSheets);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	@Test
	public void testExcelWrite(){
		file = new File("C:\\点名册.xls");
		try {
			List<ExcelSheet> excelSheets = ExcelUntil.getExcelData(file, null, 1);
			for (ExcelSheet excelSheet : excelSheets) {
				System.out.println(excelSheet);
				excelSheet.setData(new Object[][]{{1,121,12},{1,"gggg","12121"}});
				ExcelUntil.write2Excel(excelSheet, file);
			}
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
