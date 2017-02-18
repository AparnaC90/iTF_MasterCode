package com.contoso.tfsplay;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.tfs.core.TFSTeamProjectCollection;
import com.microsoft.tfs.core.clients.workitem.CoreFieldReferenceNames;
import com.microsoft.tfs.core.clients.workitem.WorkItem;
import com.microsoft.tfs.core.clients.workitem.project.Project;
import com.microsoft.tfs.core.clients.workitem.wittype.WorkItemType;

public class defectsCreation extends connectivityToTFS
{
	public static void createBugsinTFS() throws IOException
	{
		TFSTeamProjectCollection tpc=connectToTFS();
		Project project = tpc.getWorkItemClient().getProjects().get("iFT");
		System.out.println("Project");
		System.out.println(project.getName());
		WorkItemType bugType = project.getWorkItemTypes().get("Bug");
		System.out.println(bugType.getName());
		
	//	I need to fetch the date from excel and no of rows
		FileInputStream file = new FileInputStream(new File("F:\\Widget\\iTF\\RMSDefecys.xlsx")); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0); 
		
		int noofRows = sheet.getLastRowNum();
		System.out.println(noofRows);
		
		for(int i=0;i<noofRows;i++)
		{
 			sheet.getRow(i+1).getCell(2).setCellType(Cell.CELL_TYPE_STRING);
			String Priority = sheet.getRow(i+1).getCell(2).getStringCellValue();
			System.out.println(Priority);
			String Severity = sheet.getRow(i+1).getCell(1).getStringCellValue();
			System.out.println(Severity);
		//	String TestSteps = sheet.getRow(i+1).getCell(2).getStringCellValue();
		//	String Expectedresult = sheet.getRow(i+1).getCell(12).getStringCellValue();
		//	String ActualResult = sheet.getRow(i+1).getCell(13).getStringCellValue();
			String Title = sheet.getRow(i+1).getCell(0).getStringCellValue();
			System.out.println(Title);
			String createdby = sheet.getRow(i+1).getCell(3).getStringCellValue();
			System.out.println(createdby);
			
			WorkItem newWorkItem = project.getWorkItemClient().newWorkItem(bugType);
			newWorkItem.setTitle(Title);
			newWorkItem.getFields().getField(CoreFieldReferenceNames.AREA_PATH).setValue("iFT");
		//	System.out.println("Area Path");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.ASSIGNED_TO).setValue("Vaidhya P.");
		//	System.out.println("ASSIGNED TO");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.ITERATION_PATH).setValue("iFT");
		//	System.out.println("Iteration path");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.STATE).setValue("Active");
		//	System.out.println("State");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.DESCRIPTION).setValue("Login to CSMS\n Go to Dashboard\n Click Add Application\n");
		//	System.out.println("Description");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.HISTORY).setValue("Test History");
			newWorkItem.getFields().getField("Severity").setValue(Severity);
			newWorkItem.getFields().getField("Priority").setValue(Priority);
			newWorkItem.getFields().getField("Created By").setValue(createdby);
	//		newWorkItem.getFields().getField("Repro Steps").setValue(TestSteps+"\n"+Expectedresult+"\n"+ActualResult);
			
			newWorkItem.save();
			System.out.println(newWorkItem.getID());
			
		}
	//System.out.println("End of for loop");
	}
	public static void main(final String[] args) throws IOException   
	{
		createBugsinTFS();
	}
	
}
