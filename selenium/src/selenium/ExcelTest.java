package selenium;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;


public class ExcelTest {

	public static void main(String[] args) {

		/*
		String[][] arrayExcelData=getExcelData("D:\\Inputs.xls","Sheet1");

		for (int i = 0; i < (arrayExcelData.length-1); i++) {
			for (int j = 0; j < arrayExcelData[i].length; j++) {
				System.out.println(arrayExcelData[i][j]);
			}	
		}

		writeExcelData("Sheet2");*/
	}

	public static String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);

			int totalCols = sh.getColumns();
			int totalRows = sh.getRows();

			arrayExcelData = new String[totalRows][totalCols];

			for (int i= 1 ; i < totalRows; i++) {
				for (int j=0; j < totalCols; j++) {
					arrayExcelData[i-1][j] = sh.getCell(j, i).getContents();
					//System.err.println("arrayExcelData["+(i-1)+"]["+j+"]="+arrayExcelData[i-1][j]);
				}
				//System.out.println();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

	private static void writeExcelData(String sheetName) {
		// TODO Auto-generated method stub
		try {
			FileOutputStream os=new FileOutputStream("D:\\Result.xls");
			WritableWorkbook workbook = Workbook.createWorkbook(os);
			WritableSheet sheet = workbook.createSheet(sheetName, 0);
			Label caseNum = new Label(0, 5 ,"here"); 
			sheet.addCell(caseNum); 

			//Write and close the workbook
			workbook.write();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}
	private static void writeExcelData(String fileName,String sheetName) {
		// TODO Auto-generated method stub
		try {
			java.io.File fs=new java.io.File(fileName);

			WritableWorkbook workbook = Workbook.createWorkbook(fs);
			WritableSheet sheet = workbook.createSheet(sheetName, 0);

			//Create Cells with contents of different data types.
			//Also specify the Cell coordinates in the constructor

			Label label = new Label(0, 2, "A label record"); 
			sheet.addCell(label); 

			Number num=new Number(0,0,1);
			sheet.addCell(num);


			//Write and close the workbook
			workbook.write();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}
}
