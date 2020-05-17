package ITimeExport;

import java.sql.Time;

import java.util.Date;

public class ITimeEmp {
String name;
long PS_id;
Date date;
double time;




public ITimeEmp() {
	super();
	// TODO Auto-generated constructor stub
}




public ITimeEmp(String name, long pS_id, Date date, double time) {
	super();
	this.name = name;
	PS_id = pS_id;
	this.date = date;
	this.time = time;
}




public double getTime() {
	return time;
}


public void setTime(double time) {
	this.time = time;
}

public String getName() {
	return name;
}


public void setName(String name) {
	this.name = name;
}


public long getPS_id() {
	return PS_id;
}


public void setPS_id(long pS_id) {
	PS_id = pS_id;
}


public Date getDate() {
	return date;
}


public void setDate(Date date2) {
	this.date = date2;
}



}
