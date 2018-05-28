<%@include file="../checklogin.inc"%>
<%@include file="../bin/connect.inc"%>
<%@include file="../bin/getUser.inc"%>
<%@include file="../bin/connectJtrac.inc"%>
<%@include file="../bin/connectSVMJtrac.inc"%>
<%@include file="../bin/CREDOJtrac.inc"%>

<%@ page import="utils.CryptoLibrary,java.util.Date,java.text.SimpleDateFormat,java.io.*,java.util.MissingResourceException,java.util.ResourceBundle,java.util.Iterator,java.text.DecimalFormat,java.util.List,java.util.ArrayList" %>
<%@ page import="org.apache.commons.fileupload.*, org.apache.commons.fileupload.disk.DiskFileItemFactory,org.apache.commons.fileupload.DefaultFileItemFactory,org.apache.commons.fileupload.servlet.ServletFileUpload" %>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFCell, org.apache.poi.hssf.usermodel.HSSFRow, org.apache.poi.hssf.usermodel.HSSFSheet, org.apache.poi.hssf.usermodel.HSSFWorkbook, org.apache.poi.poifs.filesystem.POIFSFileSystem"%>

<%
session.removeAttribute("errorlist");
session.removeAttribute("exc");

String type="";
String projects="";

ArrayList results=new ArrayList();

String user="";
String sesId = session.getId();
String sesIdLast = (String) session.getAttribute("id");
if(sesId.equals(sesIdLast))
{
	user =  (String) session.getAttribute("user");
}
else 
{ 
	response.sendRedirect("../login.html");
	return;
}
String accuserid = (String) session.getAttribute("userid"); 
%>

<%!
	private String getCellValue(HSSFCell cell){
		DecimalFormat dc=new DecimalFormat("##############0");
        int celltype = cell.getCellType();
		String value = "";
		
		switch (celltype) {
			case HSSFCell.CELL_TYPE_BOOLEAN:
				value = String.valueOf(cell.getBooleanCellValue());
				break;                                            
			case HSSFCell.CELL_TYPE_NUMERIC:
				value = String.valueOf(dc.format(cell.getNumericCellValue()));
				break;                  
			case HSSFCell.CELL_TYPE_FORMULA:
				value = String.valueOf(dc.format(cell.getNumericCellValue()));
				break;
			case HSSFCell.CELL_TYPE_STRING:                                                
				value = String.valueOf(cell.getStringCellValue());
				break;                                            
			case HSSFCell.CELL_TYPE_BLANK:
				value = "";
				break;                                            
			case HSSFCell.CELL_TYPE_ERROR:
				value = String.valueOf(cell.getErrorCellValue());
				break;  
		}
										
        
       /* 
	    if(celltype == 0)
            value = String.valueOf((int)cell.getNumericCellValue());
        else if(celltype == 1)
            value = cell.getStringCellValue();
        else if(celltype == 3)
            value = "";
		*/		
		
		return value;		
    }
	
	private String getCellValueDate(HSSFCell cell)
	{
		int celltype = cell.getCellType();
		String value = "";
		SimpleDateFormat src = new SimpleDateFormat("MM/dd/yyyy");
		
		try
		{			
			if(getCellValue(cell).trim() != "" && getCellValue(cell).trim().length() > 0 && celltype == 0){
				int l = (int)cell.getNumericCellValue() + 68569 + 2415019;
				int no = (int)(( 4 * l ) / 146097);
				l = l - (int)((146097 * no + 3 ) / 4);
				int i = (int)(( 4000 * ( l + 1 ) ) / 1461001);
				l = l - (int)(( 1461 * i ) / 4) + 31;
				int j = (int)(( 80 * l ) / 2447);
				int nDay = l - (int)(( 2447 * j ) / 80);
				l = (int)(j / 11);
				int nMonth = j + 2 - ( 12 * l );
				int nYear = 100 * ( no - 49 ) + i + l;
				value = src.format(src.parse(nMonth+"/"+nDay+"/"+nYear));
			}
			else if(getCellValue(cell).trim() != "" && getCellValue(cell).trim().length() > 0){	
				value =	getCellValue(cell);	
			}
		}
		catch(Exception e)
		{
			System.out.println("Exception while retrieving date value ");
			e.printStackTrace();			
		}				
		return value;	
	}
	
	private boolean checkSplChars(String value){
	 	boolean result = false;
		if(value != null && value.trim() != "")
		{
			String iNum = "!@#$%^&*+=\\\';|\"<>?";
			int j=0;
			for(int i=0; i<value.length(); i++)
			{
				if (!(iNum.indexOf(value.charAt(i)) == -1)) {
					j++;
			    }
			}
			
			if(j == 0)
				result = true;
		}
        return result;
    }
	
	private boolean checkSplCharsTask(String value){
	 	boolean result = false;
		if(value != null && value.trim() != "")
		{
			String iNum = "!@#$%^&*+=\\\'-;|\"<>?";
			int j=0;
			for(int i=0; i<value.length(); i++)
			{
				if (!(iNum.indexOf(value.charAt(i)) == -1)) {
					j++;
			    }
			}
			
			if(j == 0)
				result = true;
		}
        return result;
    }
			
	private boolean validateDate(String theDate) {
		boolean result = false;
		if(theDate != null && theDate.trim() != "")
		{
			String[] dateArr = theDate.split("/");
			if(dateArr != null && dateArr.length == 3)
			{				
				if(checkNumeric(dateArr[0]) && checkNumeric(dateArr[1]) && checkNumeric(dateArr[2]))
				{
					if(Integer.parseInt(dateArr[0]) > 0 && Integer.parseInt(dateArr[0]) <= 12 && Integer.parseInt(dateArr[1]) > 0 && Integer.parseInt(dateArr[1]) <= 31 && Integer.parseInt(dateArr[2]) > 1900)			
						result = true;
				}				
			}
		}		
		return result;
	}
	
	private boolean checkNumeric(String value) {
		boolean result = false;
		if(value != null && value.trim() != "")
		{
			String iNum = "0123456789";
			int j=0;
			for(int i=0; i<value.length(); i++)
			{
				if ((iNum.indexOf(value.charAt(i)) == -1)) {
					j++;
			    }
			}
			
			if(j == 0)
				result = true;
		}
        return result;
	}
%>
<%
	String query = "";
	String finerr_result = "";
	
	String BUNDLE_NAME = "tasktype"; 
    ResourceBundle RESOURCE_BUNDLE = ResourceBundle.getBundle(BUNDLE_NAME);
    String ttype = RESOURCE_BUNDLE.getString("type");
    String[] task_type = ttype.split(",");
	
	SimpleDateFormat source = new SimpleDateFormat("MM/dd/yyyy");
	SimpleDateFormat target = new SimpleDateFormat("MM/dd/yyyy HH:mm");
	SimpleDateFormat target1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	   
	String filepath = "";//request.getParameter("taskfin");
	String path = getServletContext().getRealPath("/");
	
	try
	{
		File delfile = new File(path+"Task/ExcelUpload/");
		
		String[] filenames = delfile.list();      		
		if(filenames != null){		
			for(int j=0; j<filenames.length; j++){  
				File file1 = new File(path+"Task/ExcelUpload/"+filenames[j].trim());
				file1.delete();
			} 
		}
		
		boolean  ismultipart = FileUpload.isMultipartContent(request);
		FileItemFactory factory = new DiskFileItemFactory();
		ServletFileUpload upload = new ServletFileUpload(factory);
		List items = upload.parseRequest(request);
		
		Iterator iter = items.iterator();
        while (iter.hasNext()) 
		{             
        	FileItem item = (FileItem) iter.next();
			if(item.isFormField())
			{
				String name = item.getFieldName();
			}
			else
			{
 			    long size = item.getSize();					
				String fieldname = item.getFieldName();
				String filename = item.getName();				
				String contenttype = item.getContentType();			
				
				String fname = "";
				if(filename.lastIndexOf("\\") != -1)				
					fname = filename.substring(filename.lastIndexOf("\\"),filename.length());									
				else
					fname = filename.trim();
														
				filepath = path+"Task/ExcelUpload/"+fname;
				File uploadedFile = new File(path+"Task/ExcelUpload/"+fname);
				item.write(uploadedFile);
		   }
		}
	}
	catch(Exception e)
	{
		System.out.println("Exception in uploading");
		e.printStackTrace();
		session.setAttribute("exc","File Writing -- "+e.getMessage());
		response.sendRedirect("Success.jsp?type=exception~~");	
	}
	

	File file1 = new File(filepath.trim()); 
	if(filepath != null && file1.exists())
	{	
		String result = "";
        HSSFWorkbook wb = null;
        HSSFSheet sheet = null;
        HSSFRow row = null;
        HSSFCell cell = null;
        int lastRowNum;
        int lastCellNum;
        int firstCellNum;
		
		String fileformat = filepath.substring(filepath.lastIndexOf("."), filepath.length());                       		
            if(fileformat.trim().equalsIgnoreCase(".xls"))
			{                
                try
				{					
                    FileInputStream fis = new FileInputStream(file1);
                    POIFSFileSystem fs = new POIFSFileSystem(fis);
					
					wb = new HSSFWorkbook(fs);                    
                    sheet = wb.getSheetAt(0);  
					
					int a = 0;
                    String cellvalue = "";
										
                    int lastrownum = sheet.getLastRowNum();
					
					for(int i=5; i<=lastrownum; i++)
					{
					
						String taskname = "", projectname = "", modulename = "", tasktype = "", taskCategory="", refno = "", startdate = "", starthrs = "", startmin = "", enddate = "", endhrs = "", endmin = "", reqhrs = "", reqmin = "", status = "", comppercent = "", priority = "", resourcename = "", taskreviewer = "", taskdesc = "", err_result = "";
						int datect = 0;
					    row = sheet.getRow(i);
						
						if(row != null)
						{
							int cct = 0;
							for(a=row.getFirstCellNum(); a<=row.getLastCellNum(); a++)
							{
								cell = row.getCell((short)a);
								if(cell != null)
								{
									cellvalue = getCellValue(cell);									
									if(cellvalue.trim() != "" && cellvalue.trim().length() > 0)
										cct++;
								}	
							}
														
							if(cct > 0){
							for(a=row.getFirstCellNum(); a<=row.getLastCellNum(); a++)
							{
								cell = row.getCell((short)a);							
								if(a==1)  // Task Name									
								{

									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										//System.out.println("cellvaluecellvaluecellvaluecellvalue"+cellvalue);
										if(cellvalue != "")
										{
											if(checkSplCharsTask(cellvalue.trim()))
											{	
												taskname = cellvalue.trim();
											}
											else
											err_result += "Task Name in Row "+(i+1)+" should not contain special characters ~~";
										}
										else 
											err_result += "Task Name in Row "+(i+1)+" is mandatory ~~";
									}
									else 
										err_result += "Task Name in Row "+(i+1)+" is mandatory ~~";
								}										
								if(a==2)  // Project Name									
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										
										
										if(cellvalue != "")
										{
											if(checkSplChars(cellvalue.trim()))
											{
												for(int k =0; k<cellvalue.length(); k++)
												{
													if((int)cellvalue.charAt(k) == 8211)
														cellvalue = cellvalue.substring(0,k) + "-" + cellvalue.substring(k+1, cellvalue.length());
												}
												
												query = "select projectname from tm_projectlist where projectname = '" + cellvalue.trim() + "' and projectstatus <> 'D'";												
												rs = stmt.executeQuery(query);
												if (rs.last()) 
													projectname = cellvalue.trim();																							
												else
													err_result += "Project Name in Row "+(i+1)+" does not exist ~~";	
											}
							               else
										err_result += "Project Name in Row "+(i+1)+" should not contain special characters ~~";
										}
										else 
											err_result += "Project Name in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Project Name in Row "+(i+1)+" is mandatory ~~";
								}
								
								if(a==3)  // Module Name									
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkSplChars(cellvalue.trim()))
											{
												query = "select modulename from tm_modulelist where modulename = '" + cellvalue.trim() + "' and projectname = '" + projectname.trim() + "' and modulestatus <> 'D'";
												rs = stmt.executeQuery(query);
												if (rs.next()) 
													modulename = cellvalue.trim();																							
												else
													err_result += "Module Name in Row "+(i+1)+" does not exist ~~";	
											}
											else
												err_result += "Module Name in Row "+(i+1)+" should not contain special characters ~~";
										}
										else 
											err_result += "Module Name in Row "+(i+1)+" is mandatory ~~";
									}
									else 
										err_result += "Module Name in Row "+(i+1)+" is mandatory ~~";
								}	
											
								if(a==4)  // Task Type								
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkSplChars(cellvalue.trim()))
											{	
												for(int ii=0; ii<task_type.length; ii++)
												{		
													if(task_type[ii].trim().equalsIgnoreCase(cellvalue.trim()))														
														tasktype = cellvalue.trim();										
												}												
												if(tasktype == "")											
													err_result += "Task Type in Row "+(i+1)+" does not exist ~~";
											}
											else
												err_result += "Task Type in Row "+(i+1)+" should not contain special characters ~~";
										}
										else 
											err_result += "Task Type in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Task Type in Row "+(i+1)+" is mandatory ~~";
								}
								
								if(a==5)  // Task Category								
								{
								
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
										if(cellvalue.equals("SRS") || cellvalue.equals("CR-Development") || cellvalue.equals("CR-Testing") || cellvalue.equals("CR-Bug Fixing") || cellvalue.equals("CR-Bug Fixing Testing") || cellvalue.equals("Support Development") || cellvalue.equals("Support Testing") || cellvalue.equals("SVM-Internal") || cellvalue.equals("SVM-KT") || cellvalue.equals("SVM-Training")|| cellvalue.equals("SVM-Meeting") || cellvalue.equals("Documentation")|| cellvalue.equals("Deployment-Release"))
										{
										    taskCategory = cellvalue.trim();
												
										}		
			else
			{
			err_result += "Task Category in Row "+(i+1)+" is not valid ~~";
			
			}
												
					   }
										else 
									err_result += "Task Category in Row "+(i+1)+" is mandatory ~~";
									}
									else 
								err_result += "Task Category in Row "+(i+1)+" is mandatory ~~";
	                      }
	
// Modified by vasanthakumar.r.   purpose : To avoid ref no conflict between jtrac and eIntranet.							



if(taskCategory!="")
{	
				
if(taskCategory.substring(0,2).equals("CR"))
{
String esql = "select distinct CrEnhancement from tm_projectlist(nolock) where ProjectName='"+projectname+"'";
rs = stmt.executeQuery(esql);
String crnames1="";
String crnames=""; 
while(rs.next())
{
if(crnames1.length()>0)
{
crnames1+=",'"+rs.getString("CrEnhancement")+"'";
}
else
{
crnames1+="'"+ rs.getString("CrEnhancement")+"'";
}
}


String crnamesval=crnames1.replace("'","");
String cr[] = crnamesval.split(",");

for(int ci = 0;ci<cr.length; ci++)
{
if(crnames.length()>0)
{
crnames+=",'"+cr[ci]+"'";
}
else
{
crnames+="'"+cr[ci]+"'";
}
}



String vsql = "select distinct ss.prefix_code,im.sequence_num,ss.id,im.space_id,prefix_code,im.summary from items(nolock) im inner join spaces(nolock) ss on ss.id=im.space_id and status<>99 and ss.prefix_code in("+crnames+")";
jrs = jstmt.executeQuery(vsql);

while(jrs.next())
{
String prefixsequence=jrs.getString("prefix_code")+"-"+jrs.getString("sequence_num");

if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}



String svsql = " select distinct ss.prefix_code,im.sequence_num,ss.id,im.space_id,prefix_code,im.summary from items(nolock) im inner join spaces(nolock) ss on ss.id=im.space_id and status<>99 and ss.prefix_code in("+crnames+")";
svmrs = svmstmt.executeQuery(svsql);

while(svmrs.next())
{
String prefixsequence=svmrs.getString("prefix_code")+"-"+svmrs.getString("sequence_num");

if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}


String csql = " select distinct ss.prefix_code,im.sequence_num,ss.id,im.space_id,prefix_code,im.summary from items(nolock) im inner join spaces(nolock) ss on ss.id=im.space_id and status<>99 and  ss.prefix_code in("+crnames+")";
cdrs = cdstmt.executeQuery(csql);

while(cdrs.next())
{
String prefixsequence=cdrs.getString("prefix_code")+"-"+cdrs.getString("sequence_num");
if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}


rs.close();
jrs.close();
svmrs.close();
cdrs.close();
}

if(taskCategory.substring(0,3).equals("Sup"))
{

String esql = "select distinct ClientUat from tm_projectlist(nolock) where ProjectName='"+projectname+"'";
rs = stmt.executeQuery(esql);
String uatnames="";
String uatnames1="";

while(rs.next())
{
if(uatnames1.length()>0)
{
uatnames1+=",'"+rs.getString("ClientUat")+"'";
}
else
{
uatnames1+="'"+ rs.getString("ClientUat")+"'";
}
}


String uatnamess=uatnames1.replace("'","");
String uat[] = uatnamess.split(",");

for(int ui = 0;ui<uat.length; ui++)
{
if(uatnames.length()>0)
{
uatnames+=",'"+uat[ui]+"'";
}
else
{
uatnames+="'"+uat[ui]+"'";
}
}




String vsql = " select distinct ss.prefix_code,im.sequence_num,ss.id,im.space_id,prefix_code,im.summary from items(nolock) im inner join spaces(nolock) ss on ss.id=im.space_id and status<>99 and ss.prefix_code in("+uatnames+")";
jrs = jstmt.executeQuery(vsql);

while(jrs.next())
{
String prefixsequence=jrs.getString("prefix_code")+"-"+jrs.getString("sequence_num");

if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}


String svsql = "select distinct ss.prefix_code,im.sequence_num,ss.id,im.space_id,prefix_code,im.summary from items(nolock) im inner join spaces(nolock) ss on ss.id=im.space_id and status<>99 and ss.prefix_code in("+uatnames+")";
svmrs = svmstmt.executeQuery(svsql);

while(svmrs.next())
{
String prefixsequence=svmrs.getString("prefix_code")+"-"+svmrs.getString("sequence_num");

if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}

String csql = " select distinct ss.prefix_code,im.sequence_num,ss.id,im.space_id,prefix_code,im.summary from items(nolock) im inner join spaces(nolock) ss on ss.id=im.space_id and status<>99 and ss.prefix_code in("+uatnames+")";
System.out.println("...."+csql);
cdrs = cdstmt.executeQuery(csql);

while(cdrs.next())
{
String prefixsequence=cdrs.getString("prefix_code")+"-"+cdrs.getString("sequence_num");

if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}

rs.close();
jrs.close();
svmrs.close();
cdrs.close();
}

if(taskCategory.substring(0,3).equals("SRS"))
{
	
String ssql = "select projectcode from tm_projectlist(nolock) where ProjectName='"+projectname+"'";
rs = stmt.executeQuery(ssql);

while(rs.next())
{
String prefixsequence=rs.getString("projectcode");

if(prefixsequence.length()>0)
{ 
if(!results.contains(prefixsequence))
{
results.add(prefixsequence);
}
else
{
continue;
}
}
}
rs.close();
}
}

 if(a==6)  // Task Type Ref. No.
 {
 results.add("SRS");



   if(cell != null)
	{
	  cellvalue = getCellValue(cell);
	
	   cellvalue =(cellvalue !=null && (!"".equalsIgnoreCase(cellvalue)))?cellvalue.trim() :"";
	     String cellkey= cellvalue.substring(0, 6);
		 String taskcatkey=taskCategory.substring(0, 2);
		 String taskcatkey1=cellvalue.substring(0, 3);
		if(cellvalue.contains("SPMRDI")||cellvalue.contains("CRMRDI")){
			if(results.size()>0)
             {
				
		   if(cellvalue.length()>0)
            { 
		   
                        if(checkSplChars(cellvalue.trim())){
							if(taskCategory.contains("SRS")&&cellkey.contains("SPMRDI")){
								 refno = cellvalue.trim();
							
		                 }else if(taskcatkey.contains("CR")&&taskcatkey1.contains("CRM")){
						 
						  refno = cellvalue.trim();
						 
						 }else{
							 err_result += "Reference No. in Row "+(i+1)+" is not matching the taskcategory ~~";
						 
						 }
							 
						 }else{
                         err_result += "Reference No. in Row "+(i+1)+" should not contain special characters ~~";
            }}
            else
            {
                    err_result += "Reference No. in Row "+(i+1)+" is mandatory ~~";
            }

        }
				else if(results.contains(cellvalue))
				{
					if(cellvalue.length()>0)
                   {	    
				      if(checkSplChars(cellvalue.trim()))
					
					  refno = cellvalue.trim();
					
					  else
		       err_result += "Reference No. in Row "+(i+1)+" should not contain special characters ~~";
				  }
				 else
                 {
				  err_result += "Reference No. in Row "+(i+1)+" is mandatory ~~";
			     }
             }
			
			 else
             {
		     err_result +="Reference No. in Row "+(i+1)+" does not match in jTrac ~~";
			 }
										  
    }
}
      else
	  {
		err_result += "Reference No. in Row "+(i+1)+" is mandatory ~~";
	  }

 }

								if(a==7)  // Start Date							
								{									
									if(cell != null)
									{
										cellvalue = getCellValueDate(cell);										
										if(cellvalue != "")
										{
											if(validateDate(cellvalue.trim()))									
												startdate = cellvalue.trim();
											else{
												err_result += "Invalid Start Date Format in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											err_result += "Start Date in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Start Date in Row "+(i+1)+" is mandatory ~~";
								}	
								if(a==8)  // Start Hrs							
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkNumeric(cellvalue.trim()))	
											{
												if(Integer.parseInt(cellvalue.trim()) < 0 || Integer.parseInt(cellvalue.trim()) > 23){
													err_result += "Invalid Start Hrs in Row "+(i+1)+" ~~";	
													datect++;
												}
												else
													starthrs = cellvalue.trim();
											}	
											else{
												err_result += "Invalid Start Hrs in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											err_result += "Start Hrs in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Start Hrs in Row "+(i+1)+" is mandatory ~~";
								}	
								if(a==9)  // Start Min							
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkNumeric(cellvalue.trim()))
											{
												if(Integer.parseInt(cellvalue.trim()) < 0 || Integer.parseInt(cellvalue.trim()) > 59){
													err_result += "Invalid Start Min in Row "+(i+1)+" ~~";	
													datect++;
												}
												else
													startmin = cellvalue.trim();
											}	
											else{
												err_result += "Invalid Start Min in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											startmin = "00";
									}
									else 
										startmin = "00";
								}	
								if(a==10)  // End Date							
								{
									if(cell != null)
									{
										cellvalue = getCellValueDate(cell);
										if(cellvalue != "")
										{
											if(validateDate(cellvalue.trim()))									
												enddate = cellvalue.trim();
											else{
												err_result += "Invalid End Date Format in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											err_result += "End Date in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "End Date in Row "+(i+1)+" is mandatory ~~";
								}												
								if(a==11)  // End Hrs							
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkNumeric(cellvalue.trim()))	
											{
												if(Integer.parseInt(cellvalue.trim()) < 0 || Integer.parseInt(cellvalue.trim()) > 23){
													err_result += "Invalid End Hrs in Row "+(i+1)+" ~~";	
													datect++;
												}
												else
													endhrs = cellvalue.trim();
											}	
											else{
												err_result += "Invalid End Hrs in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											err_result += "End Hrs in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "End Hrs in Row "+(i+1)+" is mandatory ~~";
								}	
								if(a==12)  // End Min							
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkNumeric(cellvalue.trim()))
											{
												if(Integer.parseInt(cellvalue.trim()) < 0 || Integer.parseInt(cellvalue.trim()) > 59){
													err_result += "Invalid End Min in Row "+(i+1)+" ~~";	
													datect++;
												}
												else
													endmin = cellvalue.trim();
											}	
											else{
												err_result += "Invalid End Min in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											endmin = "00";
									}
									else 
										endmin = "00";
								}
								// Assigned Date Validation
								if(a==13 && startdate != "" && starthrs != "" && startmin != "" && enddate != "" && endhrs != "" && endmin != "" && datect == 0)
								{
									String sdate = target1.format(target.parse(startdate + " " + starthrs +":"+startmin));
									String edate = target1.format(target.parse(enddate + " " + endhrs +":"+endmin));
									query = "Select datediff(minute, '"+sdate+"', '"+edate+"') as tothrs";
									rs = stmt.executeQuery(query);
									while(rs.next())
									{
										if(Integer.parseInt(rs.getString("tothrs")) < 0){
											err_result += "End Date and Time should be greater than Start Date in Row "+(i+1)+" ~~";	
											datect++;
										}
									}
								}
								if(a==13)  // Required Hrs							
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkNumeric(cellvalue.trim()))	
											{
												if(Integer.parseInt(cellvalue.trim()) < 0){
													err_result += "Invalid Required Hrs in Row "+(i+1)+" ~~";	
													datect++;
												}
												else
													reqhrs = cellvalue.trim();
											}	
											else{
												err_result += "Invalid Required Hrs in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											err_result += "Required Hrs in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Required Hrs in Row "+(i+1)+" is mandatory ~~";
								}	
								if(a==14)  // Required Min							
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkNumeric(cellvalue.trim()))
											{
												if(Integer.parseInt(cellvalue.trim()) < 0 || Integer.parseInt(cellvalue.trim()) > 59){
													err_result += "Invalid Required Min in Row "+(i+1)+" ~~";	
													datect++;
												}
												else
													reqmin = cellvalue.trim();
											}	
											else{
												err_result += "Invalid Required Min in Row "+(i+1)+" ~~";																						
												datect++;
											}
										}
										else 
											reqmin = "00";
									}
									else 
										reqmin = "00";
								}
								// Req hrs. Validation
								if(a==15 && reqhrs != "" && reqmin != "" && startdate != "" && starthrs != "" && startmin != "" && enddate != "" && endhrs != "" && endmin != "" && datect == 0)
								{
									String sdate = target1.format(target.parse(startdate + " " + starthrs +":"+startmin));
									String edate = target1.format(target.parse(enddate + " " + endhrs +":"+endmin));
									query = "Select datediff(minute, '"+sdate+"', '"+edate+"') as tothrs";
									rs = stmt.executeQuery(query);
									while(rs.next())
									{
										int hrsreq = (Integer.parseInt(reqhrs)*60) + (Integer.parseInt(reqmin));
										if(hrsreq > Integer.parseInt(rs.getString("tothrs")))
											err_result += "Required Quantity of hours and min. exceed the time period in Row "+(i+1)+" ~~";	
									}
								}
								if(a==15)  // Status						
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(cellvalue.trim().equalsIgnoreCase("In Progress") || cellvalue.trim().equalsIgnoreCase("Not Started") || cellvalue.trim().equalsIgnoreCase("Completed"))
											{												
												if(cellvalue.trim().equalsIgnoreCase("In Progress"))
													status = "IP";
												if(cellvalue.trim().equalsIgnoreCase("Not Started"))
													status = "NS";
												if(cellvalue.trim().equalsIgnoreCase("Completed"))
													status = "CO";
											}	
											else
												err_result += "Incorrect Status in Row "+(i+1)+" ~~";																						
										}
										else 
											status = "NS";
									}
									else 
										status = "NS";									
								}
								if(a==16)  // Completed %						
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);																			
										if(cellvalue != "" && status != "" && status.trim().equalsIgnoreCase("IP"))
										{											
											if(checkNumeric(cellvalue.trim()))
												comppercent = cellvalue.trim();											
											else
												err_result += "Incorrect Completed % in Row "+(i+1)+" ~~";																						
										}
										else if(cellvalue != "" && status != "" && status.trim().equalsIgnoreCase("CO"))
											comppercent = "100";
										else 
											comppercent = "0";
									}
									else 
										comppercent = "0";																
								}	
								if(a==17)  // Priority						
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(cellvalue.trim().equalsIgnoreCase("Urgent") || cellvalue.trim().equalsIgnoreCase("Normal") || cellvalue.trim().equalsIgnoreCase("High"))
												priority = cellvalue.trim();
											else
												err_result += "Incorrect Priority in Row "+(i+1)+" ~~";																						
										}
										else 
											priority = "Normal";
									}
									else 
										priority = "Normal";
								}	
								if(a==18)  // Resource Name					
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										
										if(cellvalue != "")
										{																
											query = "Select user_firstname, user_lastname, user_code from logins(nolock) where user_status = 'A' and  (Ltrim(Rtrim(user_firstname)) + ' ' + Ltrim(Rtrim(user_lastname))) = '"+cellvalue.trim()+"'";
																		
											rs = stmt.executeQuery(query);
																						
											if(rs.last())
											{												
												rs.beforeFirst();
												while(rs.next())
												{																									
													resourcename = rs.getString("user_code").trim();	
																								
												}
											}
											else
												err_result += "Incorrect Resource Name in Row "+(i+1)+" ~~";						
										}
										else 
											err_result += "Resource Name in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Resource Name in Row "+(i+1)+" is mandatory ~~";
								}	
								if(a==19)  // Task Reviewer
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											String[] dummy_arr = cellvalue.trim().split(",");
											int dummyct = 0;
											if(dummy_arr != null)
											{
												for(int zz=0; zz<dummy_arr.length; zz++)
												{													
													query = "Select user_firstname, user_lastname, user_code from logins where user_status = 'A' and (Ltrim(Rtrim(user_firstname)) + ' ' + Ltrim(Rtrim(user_lastname))) = '"+dummy_arr[zz].trim()+"'";													
													rs = stmt.executeQuery(query);													
													if(rs.last())
													{
														rs.beforeFirst();
														while(rs.next())
														{																											
															if(taskreviewer.trim().length() == 0)
																taskreviewer = dummy_arr[zz].trim();	
															else
																taskreviewer += "," + dummy_arr[zz].trim();															
														}
													}
													else
														dummyct++;														
												}
												if(dummyct != 0)
													err_result += "Incorrect Task Reviewer in Row "+(i+1)+" ~~";															
											}																													
										}
										else 
											err_result += "Task Reviewer in Row "+(i+1)+" is mandatory ~~";
									}
									else 
											err_result += "Task Reviewer in Row "+(i+1)+" is mandatory ~~";
								}
								if(a==20)  // Task Description								
								{
									if(cell != null)
									{
										cellvalue = getCellValue(cell);
										if(cellvalue != "")
										{
											if(checkSplChars(cellvalue.trim()))											
												taskdesc = cellvalue.trim();
											else
												err_result += "Task Description in Row "+(i+1)+" should not contain special characters ~~";											
										}										
									}
								}																			
							} //  End - Column Check For loop	
														
							if(taskname.length() > 0 && projectname.length() > 0 && modulename.length() > 0 && resourcename.length() > 0)											
							{
								query = "select taskname from tm_tasklist where projectname = '"+projectname.trim()+"' and modulename = '"+modulename.trim()+"' and usercode= '"+resourcename.trim()+"' and taskname = '" + taskname.trim() + "' and taskstatus <> 'D'";
								rs = stmt.executeQuery(query);
								if(rs.last()) 
									err_result += "Task Name in Row "+(i+1)+" already exists ~~";																			
							}
														
							if(err_result.trim() == "" && err_result.length() == 0)
							{
							
								String sdate = target1.format(target.parse(startdate.trim() + " " + starthrs.trim() +":"+startmin.trim()));
								String edate = target1.format(target.parse(enddate.trim() + " " + endhrs.trim() +":"+endmin.trim()));
								String rhrs = reqhrs.trim() + "."+ reqmin.trim();
								query = "insert into tm_tasklist  (TaskName,TaskCode,TaskType,RefId,ProjectName,ModuleName,StartDate,CompDate,ReqHrs,TaskDesc,TaskStatus,Comp_percent,Priority,Two_days_alert,due_date_alert,post_alert,usercode,Task_reviewer,cr_user,cr_date, task_priority, task_relation, status,taskCategory) values('" + taskname.trim() + "','', '"+tasktype.trim()+ "','" +refno.trim()+"', '"+ projectname.trim()+ "','" + modulename.trim() + "','" +  sdate.trim() + "','" + edate.trim() + "'," + rhrs.trim() + ",'" + taskdesc.trim() + "','" + status.trim() + "','" + comppercent.trim() + "','"  +  priority.trim() + "','Y','N','N','" + resourcename.trim() + "','" +taskreviewer.trim() + "','" + accuserid.trim() + "',getDate(), 'P', '', 'A','"+taskCategory+"')";
								//System.out.println("insert query is "+query.trim());
								stmt.executeUpdate(query);
							}
							else
								finerr_result += err_result;
								
							} // End - cct check														
						}  // End - Row Null Check If Condn.						             
					} //  End - Row check For loop
								
					if(finerr_result.trim() != "" && finerr_result.length() > 0){
						session.setAttribute("errorlist", finerr_result.trim());					
						response.sendRedirect("Success.jsp?type=error~~");						
					}	
					else{
						System.out.println("Exception while Getting Result ");		
						response.sendRedirect("Success.jsp?type=success~~");						
					}					
				} // End - try
				catch(Exception e)
				{
					System.out.println("Exception while Getting task ");
					e.printStackTrace();					
					session.setAttribute("exc",e.getMessage());
					response.sendRedirect("Success.jsp?type=exception~~");					
				}
			}
	}
	
%>