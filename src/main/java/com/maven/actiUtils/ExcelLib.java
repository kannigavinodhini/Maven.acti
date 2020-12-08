package com.maven.actiUtils;


//class 'File' which can be reused anywhere .
//object should do all the pre requesting for us.
//do that with the help of constructor.
//create a object of file class which will take the string path as 
//where exactly the parameter is

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelLib {
	// create a constructor give Excel crtl+space.
	
	public static XSSFWorkbook wb;
	
	public ExcelLib() 
	
	{
		try {

			File src =new File("/src/test/resources/testdata/excelsheet.xltx");//create an obj of file class, give the path of your exl sheet.
			FileInputStream fis = new FileInputStream(src);//pass this file as an argu.
			wb=  new XSSFWorkbook(fis);// pass this fis as an argument.
			}
		catch (Exception e)// use base Exception just to avoid using the code
								// again and again.
		{
System.out.println(" unable to load data from excel file "+e.getMessage());
		}

	}
	//we will create a method here // static or non static depends upon the useage.
	//we wanted to use this method only when it is needed.
	public int getRowCount(int sheetnum){
		return wb.getSheetAt(sheetnum).getLastRowNum()+1;
	}
	public String getcellData(int sheetnum,int row,int cell){
		return wb.getSheetAt(sheetnum).getRow(row).getCell(cell).toString();
	}
}
