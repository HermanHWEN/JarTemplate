package org.test.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App 
{
    public static void main( String[] args ) throws IOException
    {
    	if (args.length!=2){
    		System.out.print("Require two variables!");
    		return;
    	}
//    	String filePath = "C:\\Users\\43903042\\Desktop\\test.xls";
//    	String newVersion="1.2";
    	
    	String filePath = args[0];
    	String newVersion=args[1];
    	
//    	verify excel file
    	if (filePath==null || filePath.equals("")){
    		System.out.print("Please input a valid excel file path");
    		return;
    	}
    	
    	if (newVersion==null || newVersion.equals("")){
			System.out.print("Please input a valid version!");
			return;
		}
    	
    	File excelFile = new File(filePath);
		if (!excelFile.exists()){
			System.out.print("file: " + filePath + " doesn't exist!");
			return;
		}
		
		String ExtensionName=getExtensionName(filePath);
		if (ExtensionName==null || !ExtensionName.startsWith("xls")){
			System.out.print("file: " + filePath + " is not a excel file!");
			return;
		}

		FileInputStream fileStream=new FileInputStream(excelFile);
		
		//read excel file 
		String filePathTmp =filePath;
		if (!ExtensionName.equalsIgnoreCase("xls")){//ge than 2007
	    	XSSFWorkbook wrk = new XSSFWorkbook(fileStream); 
			POIXMLProperties xmlProps = wrk.getProperties();    
			POIXMLProperties.CoreProperties coreProps =  xmlProps.getCoreProperties();
			String oldComments=coreProps.getDescription();
			
			oldComments=(oldComments==null)? "":oldComments;
			System.out.print("Old comments: " + oldComments + "\n");

			newVersion=genNewVersion(oldComments,newVersion);
			
			coreProps.setDescription(newVersion);
			System.out.print("New comments: " + coreProps.getDescription()+ "\n");
			         
			FileOutputStream outputStream1 = new FileOutputStream(filePathTmp);

			wrk.write(outputStream1);outputStream1.close();
			
		}else{//excel 2003

			HSSFWorkbook wrk=new HSSFWorkbook(fileStream);

			String oldComments=wrk.getSummaryInformation().getComments();
			oldComments=(oldComments==null)? "":oldComments;
			System.out.print("Old comments: " + oldComments + "\n");
			
			newVersion=genNewVersion(oldComments,newVersion);

			wrk.getSummaryInformation().setComments(newVersion);
			System.out.print("New comments: " + wrk.getSummaryInformation().getComments()+ "\n");
			FileOutputStream outputStream=null;
			
			outputStream = new FileOutputStream(filePathTmp);
			wrk.write(outputStream);outputStream.close();
		}
		
    }
    
    public static String getExtensionName(String filename) { 
        if ((filename != null) && (filename.length() > 0)) { 
            int dot = filename.lastIndexOf('.'); 
            if ((dot >-1) && (dot < (filename.length() - 1))) { 
                return filename.substring(dot + 1); 
            } 
        }
        return null;
    }
    
    //combine new version with old comments
    public static String genNewVersion(String comments,String newVersionStr){
    	String[] coms=comments.split("-");
		if (coms.length>1) {
			newVersionStr=comments.replace(coms[coms.length-1], newVersionStr);
		}
		return newVersionStr;
    }
}
