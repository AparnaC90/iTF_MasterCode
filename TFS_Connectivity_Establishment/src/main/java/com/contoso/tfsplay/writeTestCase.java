package com.contoso.tfsplay;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.tfs.core.TFSTeamProjectCollection;
import com.microsoft.tfs.core.clients.workitem.CoreFieldReferenceNames;
import com.microsoft.tfs.core.clients.workitem.WorkItem;
import com.microsoft.tfs.core.clients.workitem.project.Project;
import com.microsoft.tfs.core.clients.workitem.wittype.WorkItemType;

public class writeTestCase extends connectivityToTFS
{
	public static void createTestCase() throws IOException
	{
		TFSTeamProjectCollection tpc=connectToTFS();
		Project project = tpc.getWorkItemClient().getProjects().get("iTF");
		System.out.println("Project");
		System.out.println(project.getName());
		WorkItemType Type = project.getWorkItemTypes().get("Test Case");
		System.out.println(Type.getName());
		
	//	I need to fetch the date from excel and no of rows
		FileInputStream file = new FileInputStream(new File("F:\\Widget\\iTF\\TCDoc.xlsx")); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0); 
	
		
		for(int i=0;i<5;i++)
		{	
			//To check if the Test case ID already exists for the test cases
			int testcaseIDcol = 17;
			
			if(sheet.getRow(i+1).getCell(testcaseIDcol).getCellType() == Cell.CELL_TYPE_BLANK)
			{
			sheet.getRow(i+1).getCell(13).setCellType(Cell.CELL_TYPE_STRING);
			String Priority = sheet.getRow(i+1).getCell(13).getStringCellValue();
			String TestSteps = sheet.getRow(i+1).getCell(2).getStringCellValue();
			String Expectedresult = sheet.getRow(i+1).getCell(6).getStringCellValue();
			String ActualResult = sheet.getRow(i+1).getCell(7).getStringCellValue();
			String Title = sheet.getRow(i+1).getCell(7).getStringCellValue();
			String AssignedTo = sheet.getRow(i+1).getCell(16).getStringCellValue();
			String AutomationStatus = sheet.getRow(i+1).getCell(15).getStringCellValue();
			
			
			WorkItem newWorkItem = project.getWorkItemClient().newWorkItem(Type);
			newWorkItem.setTitle(Title);
			newWorkItem.getFields().getField(CoreFieldReferenceNames.AREA_PATH).setValue("iTF");
		//	System.out.println("Area Path");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.ASSIGNED_TO).setValue(AssignedTo);
		//	System.out.println("ASSIGNED TO");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.ITERATION_PATH).setValue("iTF\\Iteration 1");
		//	System.out.println("Iter path");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.STATE).setValue("Active");
		//	System.out.println("State");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.DESCRIPTION).setValue("Login to CSMS\n Go to Dashboard\n Click Add Application\n");
		//	System.out.println("Description");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.HISTORY).setValue("Test History");
		//	newWorkItem.getFields().getField("Severity").setValue(Severity);
			newWorkItem.getFields().getField("Priority").setValue(Priority);
			newWorkItem.getFields().getField(CoreFieldReferenceNames.DESCRIPTION).setValue(TestSteps+"\n"+Expectedresult+"\n"+ActualResult);
			newWorkItem.getFields().getField("Automation Status").setValue(AutomationStatus);
			
			newWorkItem.save();
			System.out.println(newWorkItem.getID());
			
			
			Cell testcaseID = sheet.getRow(i+1).createCell(17);
			testcaseID.setCellType(testcaseID.CELL_TYPE_NUMERIC);
			testcaseID.setCellValue(newWorkItem.getID());
			System.out.println(sheet.getRow(i+1).getCell(testcaseIDcol).getNumericCellValue());
			}
		file.close();
		FileOutputStream obj = new FileOutputStream(new File("F:\\Widget\\iTF\\TCDoc.xlsx"));
		workbook.write(obj);
		obj.close();
		{
			System.out.println("All the test cases are present in TFS");
		}
		
		}
		
		
		
		
	}
	public static void main(final String[] args) throws  IOException 
	{
		createTestCase();
	}
}
