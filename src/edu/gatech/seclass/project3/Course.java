package edu.gatech.seclass.project3;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import javax.script.Bindings;
import javax.script.ScriptContext;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;

public class Course {
	int numStudents;
	int numAssignments;
	int numProjects;
	HashSet<Student> studentsRoster;
	int numStudent;
	Grades grades;
	String db;
	XSSFSheet studentsSheet;
	XSSFSheet teamsSheet;
	XSSFSheet attendanceSheet;
	XSSFSheet individualGradesSheet;
	XSSFSheet individualContribsSheet;
	XSSFSheet teamGradesSheet;
	String formula = "ATT * 0.2 + AVGA * 0.4 + AVGP * 0.4"; 
	
	public Course(String db) {
		this.db = db;
		
		studentsRoster = new HashSet<Student>();
		FileInputStream file;
		try {
			file = new FileInputStream(new File(db));
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				//Get first sheet from the workbook
				studentsSheet = workbook.getSheetAt(0);
				teamsSheet = workbook.getSheetAt(1);
				attendanceSheet = workbook.getSheetAt(2);
				individualGradesSheet = workbook.getSheetAt(3);
				individualContribsSheet = workbook.getSheetAt(4);
				teamGradesSheet = workbook.getSheetAt(5);
				
				//STUDENTS SHEET
				runStudentSheet(studentsSheet);
			    
			    //TEAM SHEET
			    runTeamSheet(teamsSheet);
			    
			    //ATTENDANCE SHEET
			    runAttendanceSheet(attendanceSheet);
			    
			    //ASSIGNMENT SHEET
			    runAssignmentSheet(individualGradesSheet);
			    
			    //PROJECT SHEET
			    runProjectSheet(individualContribsSheet);
			    
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
	

	public void runStudentSheet(XSSFSheet studentsSheet){
		//STUDENTS SHEET
	    Iterator<Row> rowIterator = studentsSheet.iterator();
	    rowIterator.next();
	    while(rowIterator.hasNext()) {
	    	Row row = rowIterator.next();
	    	Student tempStudent = new Student();
	    	//Name, GTID, Email, C, C++,Java,CSJOBEX 순으로 저장됨 
	    	ArrayList tempInfoList = new ArrayList();
	    	
	        Iterator<Cell> cellIterator = row.cellIterator();
	        while(cellIterator.hasNext()) {
	             
	            Cell cell = cellIterator.next();
	            
	            switch(cell.getCellType()) {
	                case Cell.CELL_TYPE_NUMERIC:
	                	Double gtid = cell.getNumericCellValue();
	                	tempInfoList.add(gtid.intValue());
	                	//System.out.println(cell.getNumericCellValue());
	                    break;
	                case Cell.CELL_TYPE_STRING:
	                	tempInfoList.add(cell.getStringCellValue());
	                	//System.out.println(cell.getStringCellValue());
	                    break;
	            }
	        }
	        tempStudent.setName(tempInfoList.get(0).toString());
	        tempStudent.setGtid(Integer.parseInt(tempInfoList.get(1).toString()));
	        tempStudent.setEmail(tempInfoList.get(2).toString());
	        //System.out.println(tempInfoList.get(1).toString());
	        studentsRoster.add(tempStudent);
	    }
	}
	
	public void runTeamSheet(XSSFSheet teamsSheet){
	    Iterator<Row> teamRowIterator = teamsSheet.iterator();
	    teamRowIterator.next();
	    while(teamRowIterator.hasNext()) {
	    	Row row = teamRowIterator.next();
	    	//Student tempStudent = new Student();
	    	ArrayList tempInfoList = new ArrayList();
	    	String teamNumber = "";
	        Iterator<Cell> cellIterator = row.cellIterator();
	        while(cellIterator.hasNext()) {
	            Cell cell = cellIterator.next();
	            switch(cell.getCellType()) {
	            	case Cell.CELL_TYPE_STRING:
	            		teamNumber = cell.getStringCellValue();
	            		break;
	                case Cell.CELL_TYPE_NUMERIC:
	                	Double gtid = cell.getNumericCellValue();
	                	int intGtid = gtid.intValue();
	                	tempInfoList.add(intGtid);	                	
	                    break;
	            }
	        }
	        for (int i = 0; i < tempInfoList.size(); i++){
	            Student st = getStudentByID(Integer.parseInt(tempInfoList.get(i).toString()));
	            studentsRoster.remove(st);
	        	st.setTeam(teamNumber);
	        	studentsRoster.add(st);
	        }

	    }
	}
	
	public void runAttendanceSheet(XSSFSheet attendanceSheet){
	    Iterator<Row> teamRowIterator = attendanceSheet.iterator();
	    teamRowIterator.next();
	    while(teamRowIterator.hasNext()) {
	    	Row row = teamRowIterator.next();
	    	ArrayList attendanceList = new ArrayList();
	    	
	        Iterator<Cell> cellIterator = row.cellIterator();
	        while(cellIterator.hasNext()) {
	            Cell cell = cellIterator.next();
	            
	            switch(cell.getCellType()) {
	                case Cell.CELL_TYPE_NUMERIC:
	                	Double data = cell.getNumericCellValue();
	                	int intData = data.intValue();
	                	attendanceList.add(intData);	                	
	                    break;
	            }
	        }
            Student st = getStudentByID(Integer.parseInt(attendanceList.get(0).toString()));
            studentsRoster.remove(st);
        	st.setAttendance(Integer.parseInt(attendanceList.get(1).toString()));
        	studentsRoster.add(st);
	        

	    }
	}
	
	public void runAssignmentSheet(XSSFSheet individualGradesSheet) {
		// TODO Auto-generated method stub
		Iterator<Row> teamRowIterator = individualGradesSheet.iterator();
	    Row headerRow = teamRowIterator.next();
	    Iterator<Cell> cellIter = headerRow.cellIterator();
	    ArrayList list = new ArrayList();
	    while(cellIter.hasNext()) {
        	Cell cell = cellIter.next();
            switch(cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                	list.add(cell.getStringCellValue());	                	
                    break;
            }
        }
	    this.numAssignments = list.size()-1;
	    
	    ArrayList assignmentList = new ArrayList();
	    while(teamRowIterator.hasNext()) {
	    	Row row = teamRowIterator.next();
	    	Iterator<Cell> cellIterator = row.cellIterator();
	    	
	        while(cellIterator.hasNext()) {
	        	Cell cell = cellIterator.next();
	            switch(cell.getCellType()) {
	                case Cell.CELL_TYPE_NUMERIC:
	                	Double data = cell.getNumericCellValue();
	                	int intData = data.intValue();
	                	assignmentList.add(intData);	                	
	                    break;
	            }
	        }
	    }
	    
	}
	
	public void runProjectSheet(XSSFSheet individualContribsSheet){
		Iterator<Row> teamRowIterator = individualContribsSheet.iterator();
	    Row headerRow = teamRowIterator.next();
	    Iterator<Cell> cellIter = headerRow.cellIterator();
	    ArrayList list = new ArrayList();
	    while(cellIter.hasNext()) {
        	Cell cell = cellIter.next();
            switch(cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                	list.add(cell.getStringCellValue());	                	
                    break;
            }
        }
	    this.numProjects = list.size()-1;
	}
	
	public int getNumStudents(){
		return studentsRoster.size();
	}
	
	public int getNumAssignments(){
		//return this.numAssignments/studentsRoster.size() -1 ;
		return this.numAssignments;
	}
	
	public int getNumProjects(){
		return this.numProjects;
	}
	
	public HashSet<Student> getStudents(){
		return studentsRoster;
	}
	
	public Student getStudentByName(String name){
		Iterator<Student> studentIterator = studentsRoster.iterator();
	    while(studentIterator.hasNext()) {
             
            Student st = studentIterator.next();
            if (st.getName().equals(name)) return st;
        }
		Student student = new Student();
		return student;
	}
	
	public Student getStudentByID(int id){
		Iterator<Student> studentIterator = studentsRoster.iterator();
	    while(studentIterator.hasNext()) {
             
            Student st = studentIterator.next();
            if (st.getGtid() == id) return st;
        }
		Student student = new Student();
		return student;

	}


	public void addAssignment(String newAssignment) {
		// TODO Auto-generated method stub
		FileInputStream file;
		try {
			file = new FileInputStream(new File(db));
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				//Get first sheet from the workbook

				individualGradesSheet = workbook.getSheetAt(3);
				
				int indForNewAssignment = this.numAssignments + 1;
				
				Iterator<Row> teamRowIterator = individualGradesSheet.iterator();
			    Row headerRow = teamRowIterator.next();
			    headerRow.createCell(indForNewAssignment).setCellValue(newAssignment);
			    
//			    
//				
//			    Cell cell = null;
//			    cell=individualGradesSheet.getRow(0).getCell(indForNewAssignment);
//			    cell.setCellValue(newAssignment);
			    
			    FileOutputStream outFile =new FileOutputStream(new File(db));
			    workbook.write(outFile);
			    outFile.close();
			    
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		
		
	    //this.numAssignments ++;
	    
	    
	}


	public void updateGrades(Grades grades) {
		// TODO Auto-generated method stub		
		FileInputStream file;
		try {
			file = new FileInputStream(new File(db));
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				//Get first sheet from the workbook

				individualGradesSheet = workbook.getSheetAt(3);
				individualContribsSheet = workbook.getSheetAt(4);
				teamGradesSheet = workbook.getSheetAt(5);
			    
			    //ASSIGNMENT SHEET
			    runAssignmentSheet(individualGradesSheet);
			    
			    //PROJECT SHEET
			    runProjectSheet(individualContribsSheet);
			    
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


	public void addGradesForAssignment(String assignmentName, HashMap<Student, Integer> grades) {
		FileInputStream file;
		try {
			file = new FileInputStream(new File(db));
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				individualGradesSheet = workbook.getSheetAt(3);
				
				Iterator<Row> teamRowIterator = individualGradesSheet.iterator();
			    Row headerRow = teamRowIterator.next();
			    
			    /////Find which column, which assignment to change
			    Iterator<Cell> cellIter = headerRow.cellIterator();
			    int whichCol = 0;
			    while(cellIter.hasNext()) {
		        	Cell cell = cellIter.next();
		            if(cell.getStringCellValue().equals(assignmentName)){
		            	break;
		            }
		            whichCol++;
		        }
			    
			    ///// Retrieve all grades to HashSet<Integer, Integer>
			    HashMap<Integer, Integer> idToGradeMap = new HashMap<Integer, Integer>();
			    Iterator<Student> students = grades.keySet().iterator();
		        while( students.hasNext() ){
		            Student st = students.next();
		            idToGradeMap.put(st.getGtid(), grades.get(st).intValue());
		        }
			    
			    
			    ///// Insert or update grades for corresponding students
			    int whichRow = 1;
			    while(teamRowIterator.hasNext()) {
			    	Row row = teamRowIterator.next();
			    	
			    	Double gtid = row.getCell(0).getNumericCellValue();
                	int currentGtid = gtid.intValue();
                	//System.out.println("Current GTID: " + currentGtid);                	
                	
                	if(idToGradeMap.containsKey(currentGtid)){
                		int grade = idToGradeMap.get(currentGtid);
                    	//System.out.println("Current grade: " + grade);

                		row.createCell(whichCol).setCellValue(grade);
                	
                	}

			    }
			    
			    ///////////////////////////////////
			    ////Only for Test
			    FileOutputStream outFile =new FileOutputStream(new File(db));
			    workbook.write(outFile);
			    outFile.close();
			    
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}


	public int getAverageAssignmentsGrade(Student student) {
		// TODO Auto-generated method stub
		double average = 0;
		FileInputStream file;
		ArrayList gradeList = new ArrayList();;
		try {
			file = new FileInputStream(new File(db));
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				individualGradesSheet = workbook.getSheetAt(3);
				
				Iterator<Row> assignmentRowIterator = individualGradesSheet.iterator();
			    Row headerRow = assignmentRowIterator.next();

			    while(assignmentRowIterator.hasNext()) {
			    	Row row = assignmentRowIterator.next();
			    	
			    	Double gtid = row.getCell(0).getNumericCellValue();
                	int currentGtid = gtid.intValue();
                	//System.out.println("Current GTID: " + currentGtid);                	
                	gradeList = new ArrayList();
                	if(student.getGtid()==currentGtid){
                	    Iterator<Cell> cellIter = row.cellIterator();
                	    Cell cell = cellIter.next();
                	    while(cellIter.hasNext()) {
                        	cell = cellIter.next();
                            switch(cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                	gradeList.add(cell.getNumericCellValue());	                	
                                    break;
                            }
                        }
                	    break;
                	}

			    }
			    
			    //Calculate average value in gradeList
			    
			    double sum = 0;
			    for (int i = 0 ; i < gradeList.size(); i++){
			    	sum = sum + Double.parseDouble(gradeList.get(i).toString());
			    	//System.out.println(gradeList.get(i));
			    }
			    average = sum / gradeList.size();
			    
			    ///////////////////////////////////
			    ////Only for Test
			    FileOutputStream outFile =new FileOutputStream(new File(db));
			    workbook.write(outFile);
			    outFile.close();
			    
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return (int) Math.round(average);
	}


	public int getAverageProjectsGrade(Student student) {
		Student newStudent = new Student();
		Iterator<Student> roster = studentsRoster.iterator();
        while( roster.hasNext() ){
            Student st = roster.next();
            if(st.getGtid() == student.getGtid()){
            	newStudent = st;
            }
        }
        
        String teamName = newStudent.getTeam();
		//System.out.println(newStudent.getTeam());		

        
		double average = 0;
		FileInputStream file;
		ArrayList gradeList = new ArrayList();
		ArrayList teamGradeList = new ArrayList();
		try {
			file = new FileInputStream(new File(db));
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				individualContribsSheet = workbook.getSheetAt(4);
				teamGradesSheet = workbook.getSheetAt(5);
				Iterator<Row> teamGradeRowIterator = teamGradesSheet.iterator();
				teamGradeRowIterator.next();
				
				Iterator<Row> projectRowIterator = individualContribsSheet.iterator();
			    Row headerRow = projectRowIterator.next();

			    while(projectRowIterator.hasNext()) {
			    	Row row = projectRowIterator.next();
			    	
			    	Double gtid = row.getCell(0).getNumericCellValue();
                	int currentGtid = gtid.intValue();
                	//System.out.println("Current GTID: " + currentGtid);                	
                	gradeList = new ArrayList();
                	if(student.getGtid()==currentGtid){
                	    Iterator<Cell> cellIter = row.cellIterator();
                	    Cell cell = cellIter.next();
                	    //System.out.println(cell.getNumericCellValue());
                	    while(cellIter.hasNext()) {
                        	cell = cellIter.next();
                            switch(cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                	//System.out.println(cell.getNumericCellValue());
                                	gradeList.add(cell.getNumericCellValue());	                	
                                    break;
                            }
                        }
                	    break;
                	}
			    }
			    //Loop through team grade sheet
			    while(teamGradeRowIterator.hasNext()) {
			    	Row row = teamGradeRowIterator.next();
			    	
			    	String team = row.getCell(0).getStringCellValue();
                	teamGradeList = new ArrayList();
                	if(team.equals(teamName)){
                	    Iterator<Cell> cellIter = row.cellIterator();
                	    Cell cell = cellIter.next();
                	    //System.out.println(cell.getStringCellValue());
                	    while(cellIter.hasNext()) {
                        	cell = cellIter.next();
                            switch(cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                	//System.out.println(cell.getNumericCellValue());
                                	teamGradeList.add(cell.getNumericCellValue());	                	
                                    break;
                            }
                        }
                	    break;
                	}
			    }
			    
			    //Calculate average value in gradeList
			    
			    double sum = 0;
			     
			    for (int i = 0 ; i < gradeList.size(); i++){
			    	double individualContribution = Double.parseDouble(gradeList.get(i).toString())*0.01;
			    	double teamGrade = Double.parseDouble(teamGradeList.get(i).toString());
			    	sum = sum + teamGrade*individualContribution;
			    	
			    	//System.out.println(teamGradeList.get(i));
			    }
			    average = sum / gradeList.size();
			    
			    ///////////////////////////////////
			    ////Only for Test
			    FileOutputStream outFile =new FileOutputStream(new File(db));
			    workbook.write(outFile);
			    outFile.close();
			    
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return (int) Math.round(average);
	}


	public void addIndividualContributions(String projectName, HashMap<Student, Integer> contributions) {
		FileInputStream file;
		try {
			file = new FileInputStream(new File(db));
			XSSFWorkbook workbook;
			try {
				workbook = new XSSFWorkbook (file);
				individualContribsSheet = workbook.getSheetAt(4);
				
				Iterator<Row> indGradeRowIterator = individualContribsSheet.iterator();
			    Row headerRow = indGradeRowIterator.next();
			    
			    /////Find which column, which assignment to change
			    Iterator<Cell> cellIter = headerRow.cellIterator();
			    int whichCol = 0;
			    while(cellIter.hasNext()) {
		        	Cell cell = cellIter.next();
		            if(cell.getStringCellValue().equals(projectName)){
		            	break;
		            }
		            whichCol++;
		        }
			    
			    ///// Retrieve all grades to HashSet<Integer, Integer>
			    HashMap<Integer, Integer> idToGradeMap = new HashMap<Integer, Integer>();
			    Iterator<Student> students = contributions.keySet().iterator();
		        while( students.hasNext() ){
		            Student st = students.next();
		            idToGradeMap.put(st.getGtid(), contributions.get(st).intValue());
		        }
			    
			    
			    ///// Insert or update grades for corresponding students
			    int whichRow = 1;
			    while(indGradeRowIterator.hasNext()) {
			    	Row row = indGradeRowIterator.next();
			    	
			    	Double gtid = row.getCell(0).getNumericCellValue();
                	int currentGtid = gtid.intValue();
                	//System.out.println("Current GTID: " + currentGtid);                	
                	
                	if(idToGradeMap.containsKey(currentGtid)){
                		int grade = idToGradeMap.get(currentGtid);
                    	//System.out.println("Current grade: " + grade);

                		row.createCell(whichCol).setCellValue(grade);
                	
                	}

			    }
			    
			    ///////////////////////////////////
			    ////Only for Test
			    FileOutputStream outFile =new FileOutputStream(new File(db));
			    workbook.write(outFile);
			    outFile.close();
			    
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}


	public int getAttendance(Student student) {
		Student target = new Student();
		
		HashMap<Integer, Integer> idToGradeMap = new HashMap<Integer, Integer>();
	    Iterator<Student> students = studentsRoster.iterator();
        while( students.hasNext() ){
            Student st = students.next();
            if(st.getGtid() == student.getGtid()){
            	target = st;
            	break;
            }
        }
        return target.getAttendance();
   }


	public String getTeam(Student student) {
		Student target = new Student();
		
		HashMap<Integer, Integer> idToGradeMap = new HashMap<Integer, Integer>();
	    Iterator<Student> students = studentsRoster.iterator();
        while( students.hasNext() ){
            Student st = students.next();
            if(st.getGtid() == student.getGtid()){
            	target = st;
            	break;
            }
        }
        return target.getTeam();
	}


	public void setFormula(String string) {
		// TODO Auto-generated method stub
		this.formula = string;
	}


	public Object getOverallGrade(Student student) throws GradeFormulaException{

        int AVGA = getAverageAssignmentsGrade(student);
        int AVGP = getAverageProjectsGrade(student);
        int ATT = getAttendance(student);
        
        //String expr = getFormula();
        String expr = getFormula();
        // Get the JavaScript engine
        ScriptEngine engine = new ScriptEngineManager().getEngineByExtension("js");
  
        // Bind variables
        Bindings bindings = engine.createBindings();
        bindings.put("AVGA", AVGA);
        bindings.put("AVGP", AVGP);
        bindings.put("ATT", ATT);
        //bindings.put("AVG", AVGA);
        
        engine.setBindings(bindings, ScriptContext.ENGINE_SCOPE);

        Object result = null;
		try {
			result = engine.eval(expr);
			//System.out.println("The expression '" + expr + "' evaluates to: " + result);
		} 
		catch (Exception e) {
			throw new GradeFormulaException("Malformed formulas given !! ");
		}  
  
        
		//return Integer.parseInt(result.toString());
		String temp = result.toString();
		double grade = Double.parseDouble(temp);
		
		return (int)Math.round(grade);
	}


	public String getFormula() {
		// TODO Auto-generated method stub
		return this.formula;
	}


	public String getEmail(Student student) {
		Student target = new Student();
		
		HashMap<Integer, Integer> idToGradeMap = new HashMap<Integer, Integer>();
	    Iterator<Student> students = studentsRoster.iterator();
        while( students.hasNext() ){
            Student st = students.next();
            if(st.getGtid() == student.getGtid()){
            	target = st;
            	break;
            }
        }
        return target.getEmail();
	}
	
}
