package com.contoso.tfsplay;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.JFrame;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.tfs.core.TFSTeamProjectCollection;
import com.microsoft.tfs.core.clients.workitem.WorkItem;
import com.microsoft.tfs.core.clients.workitem.WorkItemClient;
import com.microsoft.tfs.core.clients.workitem.query.WorkItemCollection;

public class exportAndUpdateTestcases extends connectivityToTFS
{
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
        	 DateFormat df = new SimpleDateFormat("dd/MM/yy HH:mm:ss");
        	 Date dateobj = new Date();
        	 
        	FileOutputStream outputStream = new FileOutputStream(".\\Documents\\"+dateobj+"\\Testcase.xlsx");
        			workbook.write(outputStream);
        		
        			
        }
	 catch(Exception e)
        {
		
        }
        }
}
