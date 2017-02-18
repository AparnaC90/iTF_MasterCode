package com.contoso.tfsplay;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.JTable;
import javax.swing.table.TableModel;
import javax.xml.parsers.DocumentBuilder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hsqldb.Table;
import org.hsqldb.util.CSVWriter;

import com.microsoft.tfs.core.TFSTeamProjectCollection;
import com.microsoft.tfs.core.clients.workitem.CoreFieldReferenceNames;
import com.microsoft.tfs.core.clients.workitem.WorkItem;  
import com.microsoft.tfs.core.clients.workitem.WorkItemClient;
import com.microsoft.tfs.core.clients.workitem.WorkItemStateListener;
import com.microsoft.tfs.core.clients.workitem.exceptions.UnableToSaveException;
import com.microsoft.tfs.core.clients.workitem.fields.FieldCollection;
import com.microsoft.tfs.core.clients.workitem.files.AttachmentCollection;
import com.microsoft.tfs.core.clients.workitem.link.LinkCollection;
import com.microsoft.tfs.core.clients.workitem.project.Project;

import com.microsoft.tfs.core.clients.workitem.query.WorkItemCollection;
import com.microsoft.tfs.core.clients.workitem.revision.RevisionCollection;
import com.microsoft.tfs.core.clients.workitem.wittype.WorkItemType;
import com.microsoft.tfs.core.httpclient.Credentials;
import com.microsoft.tfs.core.httpclient.UsernamePasswordCredentials;
import com.microsoft.tfs.core.httpclient.util.URIUtil;
import com.microsoft.tfs.core.util.URIUtils;

import com.microsoft.tfs.core.clients.*;

import java.awt.BorderLayout;

import javax.swing.JFrame;
import javax.swing.JScrollPane;

public class QueryWorkItemExample 
{
	
	
	public static TFSTeamProjectCollection connectToTFS() 
	{
		TFSTeamProjectCollection tpc=null;
		Credentials credentials;
		
		credentials = new UsernamePasswordCredentials("caparna","Undertakerjohn@11");
		tpc = new TFSTeamProjectCollection(URIUtils.newURI("http://10.0.10.79:8080/tfs/DefaultCollection"), credentials);
	    return tpc;
	}
	
	public static void ExportToExcel()
	{
		TFSTeamProjectCollection tpc=connectToTFS();
		System.out.println("Connected to TFS");
		 WorkItemClient workItemClient = tpc.getWorkItemClient();
		 
		 String wiqlQuery = "Select [ID], [Title], [State], [Assigned To], [Area Path], [Iteration Path] from WorkItems where [Work Item Type] = 'Test Case' AND [Area Path] UNDER 'WidgetFactory\\RACFtoAD' and [Iteration Path] UNDER 'WidgetFactory\\RACF\\RACF Phase 1\\Sprint 2' order by Title";
	
		 
			// WorkItemCollection workItems = workItemClient .query("SELECT [System.Id], [System.Title], [System.AssignedTo], [System.CreatedBy], [System.CreatedDate], [System.AreaPath], [System.IterationPath], [System.Description] FROM WorkItems WHERE [System.WorkItemType] = 'Work Item Type 1' AND [System.Title] CONTAINS 'GUI' ORDER BY [System.Id]");
		 WorkItemCollection workItems = workItemClient.query(wiqlQuery);
			 System.out.println("No of work items"+workItems.size());
			 System.out.println(workItems);
			 
		int[] Id=new int[workItems.size()];
        String[] Title=new String[workItems.size()];
        String[] Type=new String[workItems.size()];
       String[] Project =  new String[workItems.size()];
       
       WorkItem workItem = workItems.getWorkItem(1);
       System.out.println(workItem.getProject().getName());
       
        
	 
        
		/* if(i>100)
		 {
			 System.out.println("[..]");
			 break;
		 }
		 WorkItem workItem = workItems.getWorkItem(i);
		 System.out.println(workItem.getID()+"\t"+workItem.getTitle()+"\t"+workItem.getType().getName()+"\t");
		 }*/
		
        String data[][]= new String[workItems.size()][workItems.size()];
        String column[] = {"Title","Type"};
        
        try
        {
        	JFrame f;
        	 f=new JFrame(); 
        	 XSSFWorkbook workbook = new XSSFWorkbook();
        	 XSSFSheet sheet = workbook.createSheet("Test Cases");
        	 for(int i=0;i<workItems.size();i++)
        	 {
        		 System.out.println("in for loop for i="+i);
       // 	WorkItem workItem = workItems.getWorkItem(i);
        	Id[i] = workItem.getID();
        	Title[i]=workItem.getTitle();
        	Type[i]=workItem.getType().getName();
        	Project[i]=workItem.getProject().getName();
        	if(workItem.getFields().iterator().hasNext())
        	{
        		System.out.println(workItem.getFields().contains("Widget Factory"));
        	}
        	
        	Row row = sheet.createRow(i);
        	for(int j=0;j<4;j++)
        	{
        		if(j==0)
        		{
        	Cell cell=row.createCell(j);
        	cell.setCellValue(Id[i]);
        		}
        		else if(j==1)
        		{
        			Cell cell=row.createCell(j);
        			cell.setCellValue(Title[i]);
        		}
        		else if(j==2)
        		{	
        			Cell cell=row.createCell(j);
        			cell.setCellValue(Type[i]);
        		}
        	}
        		
        	System.out.println(Id[i]+"\t"+Title[i]+"\t"+Type[i]+"\t"+Project[i]+"\n");
        /*	data[i][i]={{Title[i],Type[i]}};
           
        	 }
            JTable jt=new JTable(data,column);
            jt.setBounds(30,40,200,300);
            JScrollPane sp=new JScrollPane(jt);
            f.add(sp);          
            f.setSize(300,400);    
            f.setVisible(true);*/
        	 }
        	FileOutputStream outputStream = new FileOutputStream("F:\\Widget\\iTF\\Testcase.xlsx");
        			workbook.write(outputStream);
        		
        			
        }
	 catch(Exception e)
        {
		
        }
        }
	public static void createBugsinTFS() throws IOException
	{
		TFSTeamProjectCollection tpc=connectToTFS();
		Project project = tpc.getWorkItemClient().getProjects().get("iTF");
		System.out.println("Project");
		System.out.println(project.getName());
		WorkItemType bugType = project.getWorkItemTypes().get("Bug");
		System.out.println(bugType.getName());
		
	//	I need to fetch the date from excel and no of rows
		FileInputStream file = new FileInputStream(new File("F:\\Widget\\iTF\\TC Document.xlsx")); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0); 
		
		for(int i=0;i<3;i++)
		{
			sheet.getRow(i+1).getCell(13).setCellType(Cell.CELL_TYPE_STRING);
			String Priority = sheet.getRow(i+1).getCell(13).getStringCellValue();
			String Severity = sheet.getRow(i+1).getCell(14).getStringCellValue();
			String TestSteps = sheet.getRow(i+1).getCell(2).getStringCellValue();
			String Expectedresult = sheet.getRow(i+1).getCell(6).getStringCellValue();
			String ActualResult = sheet.getRow(i+1).getCell(7).getStringCellValue();
			String Title = sheet.getRow(i+1).getCell(7).getStringCellValue();
			
			WorkItem newWorkItem = project.getWorkItemClient().newWorkItem(bugType);
			newWorkItem.setTitle(Title);
			newWorkItem.getFields().getField(CoreFieldReferenceNames.AREA_PATH).setValue("iTF");
		//	System.out.println("Area Path");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.ASSIGNED_TO).setValue("Venkateswara Reddy Gudugunur Rami");
		//	System.out.println("ASSIGNED TO");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.ITERATION_PATH).setValue("iTF\\Iteration 1");
		//	System.out.println("Iter path");
			newWorkItem.getFields().getField(CoreFieldReferenceNames.STATE).setValue("Active");
		//	System.out.println("State");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.DESCRIPTION).setValue("Login to CSMS\n Go to Dashboard\n Click Add Application\n");
		//	System.out.println("Description");
		//	newWorkItem.getFields().getField(CoreFieldReferenceNames.HISTORY).setValue("Test History");
			newWorkItem.getFields().getField("Severity").setValue(Severity);
			newWorkItem.getFields().getField("Priority").setValue(Priority);
			newWorkItem.getFields().getField("Repro Steps").setValue(TestSteps+"\n"+Expectedresult+"\n"+ActualResult);
			
			newWorkItem.save();
			System.out.println(newWorkItem.getID());
			
		}
		System.out.println("End of for loop");
	}
public static void main(final String[] args) throws IOException   
{
	ExportToExcel();
	createBugsinTFS();
}

}

	

