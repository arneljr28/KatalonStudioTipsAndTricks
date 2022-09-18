package tipsandtricks

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.googlecode.javacv.cpp.avformat.AVIndexEntry
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import groovy.json.StringEscapeUtils

import com.kms.katalon.core.testdata.reader.ExcelFactory
import com.kms.katalon.core.testdata.ExcelData

import java.util.Random

import internal.GlobalVariable

public class DefineSpecificExcelSheet {
	
	File path = new File("ExcelData//DefineSpecificExcelSheet.xlsx")
	
	public String getRandomExcelValue(String sheetName, String columnName) {
		Random random = new Random()
		
		ExcelData excelData = ExcelFactory.getExcelDataWithDefaultSheet(path.getAbsolutePath(), sheetName, true)
		
		int selectRandomItem = new Random().nextInt(excelData.getRowNumbers())
		
		return excelData.getValue(columnName, selectRandomItem)
	}
	
	public String getRandomExcelValue(String sheetName, int columnName) {
		Random random = new Random()
		
		ExcelData excelData = ExcelFactory.getExcelDataWithDefaultSheet(path.getAbsolutePath(), sheetName, true)
		
		int selectRandomItem = new Random().nextInt(excelData.getRowNumbers())
		
		return excelData.getValue(columnName, selectRandomItem)
	}
	
	public List<String> getAllDataFromSpecificSheet(String sheetName) {
		
		ExcelData excelData = ExcelFactory.getExcelDataWithDefaultSheet(path.getAbsolutePath(), sheetName, true)
		
		return excelData.getAllData()
	}
	
	public String findNameFromDefinedSheet(String sheetName, String name) {
		ExcelData excelData = ExcelFactory.getExcelDataWithDefaultSheet(path.getAbsolutePath(), sheetName, true)
		
		for (int nameIndex = 0; nameIndex < excelData.getRowNumbers(); nameIndex++) {
			
		}
		
		return ""
	}
}
