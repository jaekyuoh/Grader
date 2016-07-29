package edu.gatech.seclass.project3;

public class Student {
	String name;
	int Gtid;
	int attendance;
	String team;
	String email;
	Course course;
	
	
	public Student(){
		
	}
	public Student (String name, int Gtid){
		this.name = name;
		this.Gtid = Gtid;
	}
	public Student (String name, int Gtid,String email){
		this.name = name;
		this.Gtid = Gtid;
		this.email = email;
	}
	
	public Student(String name, int Gtid, String team, String email){
		this.name = name;
		this.Gtid = Gtid;
		this.team = team;
		this.email = email;
	}
	
	public Student(String name, int Gtid, int attendance, String team, String email){
		this.name = name;
		this.Gtid = Gtid;
		this.attendance = attendance;
		this.team = team;
		this.email = email;
	}
	
	public Student(String name, int Gtid, Course course){
		this.name = name;
		this.Gtid = Gtid;
		this.course = course;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getGtid() {
		return Gtid;
	}

	public void setGtid(int numb) {
		this.Gtid = numb;
	}

	public int getAttendance() {
		return attendance;
	}

	public void setAttendance(int attendance) {
		this.attendance = attendance;
	}

	public String getTeam() {
		return team;
	}

	public void setTeam(String team) {
		this.team = team;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public Course getCourse() {
		return course;
	}

	public void setCourse(Course course) {
		this.course = course;
	}
	
	
	
}
