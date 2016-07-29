package edu.gatech.seclass.project3;

import static org.junit.Assert.*;
import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.HashSet;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import edu.gatech.seclass.project3.Course;
import edu.gatech.seclass.project3.Grades;
import edu.gatech.seclass.project3.Student;
import edu.gatech.seclass.project3.Students;

import org.junit.Test;

public class AdditionalCourseTest {

    Students students = null;
    Grades grades = null;
    Course course = null;
    static final String DB = "DB/CourseDatabase6300.xlsx";
    static final String DB_GOLDEN = "DB/CourseDatabase6300-golden.xlsx";

    @Before
    public void setUp() throws Exception {
        FileSystem fs = FileSystems.getDefault();
        Path dbfilegolden = fs.getPath(DB_GOLDEN);
        Path dbfile = fs.getPath(DB);
        Files.copy(dbfilegolden, dbfile, REPLACE_EXISTING);
        course = new Course(DB);
    }

    @After
    public void tearDown() throws Exception {
        students = null;
        grades = null;
        course = null;
    }
    
    @Test
    public void testAddStudent() {
    	
        Student student1 = new Student("Jaekyu Oh", 901234517, "Team4", "email1@gatech.edu");
        Student student2 = new Student("JQ O", 901234518, "Team4", "email2@gatech.edu");
        
        course.addStudent(student1);
        assertEquals(17, course.getNumStudents());
        
        course.addStudent(student2);        
        assertEquals(18, course.getNumStudents());
    }
    
    @Test
    public void testAddProject() {
        course.addProject("Project: Project 4");
        course.updateGrades(new Grades(DB));
        assertEquals(4, course.getNumAssignments());
        
        course.addProject("Project: Project 5");
        course.updateGrades(new Grades(DB));
        
        assertEquals(5, course.getNumProjects());
    }
    
    @Test
    public void testGetAverageTeamProjectsGrade() {
        // Rounded to the closest integer
        Team team1 = new Team("Team 1", course);        
        assertEquals(96, course.getAverageTeamProjectsGrade(team1));
    }

    @Test
    public void testAddGradesForProject() {
        String projectName = "Project: Project4";
        Team team1 = new Team("Team 1", course);
        Team team2 = new Team("Team 2", course);
        Team team3 = new Team("Team 3", course);
        course.addProject(projectName);
        course.updateGrades(new Grades(DB));
        
        HashMap<Team, Integer> grades = new HashMap<Team, Integer>();
        grades.put(team1, 87);
        grades.put(team2, 94);
        grades.put(team3, 100);
        course.addGradesForProject(projectName, grades);
        course.updateGrades(new Grades(DB));
        
        assertEquals(94, course.getAverageTeamProjectsGrade(team1));
        assertEquals(96, course.getAverageTeamProjectsGrade(team2));
        assertEquals(95, course.getAverageTeamProjectsGrade(team3));
    }

}
