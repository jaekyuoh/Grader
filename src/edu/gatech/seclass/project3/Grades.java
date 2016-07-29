package edu.gatech.seclass.project3;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Grades {
	
	
	
	public Grades(String db){
		FileInputStream file;
		try {
			file = new FileInputStream(new File(db));
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				//Get grade sheets from the workbook

				XSSFSheet individualGradesSheet = workbook.getSheetAt(3);
				XSSFSheet individualContribsSheet = workbook.getSheetAt(4);
				XSSFSheet teamGradesSheet = workbook.getSheetAt(5);
			    
			    //ASSIGNMENT SHEET
			    runAssignmentSheet(individualGradesSheet);
			    
			    //PROJECT SHEET
			    runProjectSheet(individualContribsSheet);
			    
			    //PROJECT SHEET
			    runTeamGradeSheet(teamGradesSheet);
			    
			    file.close();
				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	



	public void runAssignmentSheet(XSSFSheet individualGradesSheet) {

	    
	}
	
	public void runProjectSheet(XSSFSheet individualContribsSheet){

	}
	
	
	private void runTeamGradeSheet(XSSFSheet teamGradesSheet) {
		// TODO Auto-generated method stub
		
	}
	

	
}
