package ITimeExport;

import java.util.Date;

public class CaPpmEmp   {
long PsId;
String name;
Date date;
double time;
String ManagerName;

public CaPpmEmp() {
	super();
	// TODO Auto-generated constructor stub
}



public CaPpmEmp(long psId, String name, Date date, double time, String managerName) {
	super();
	PsId = psId;
	this.name = name;
	this.date = date;
	this.time = time;
	ManagerName = managerName;
}



public double getTime() {
	return time;
}



public void setTime(double time) {
	this.time = time;
}



public long getPsId() {
	return PsId;
}

public void setPsId(long psId) {
	PsId = psId;
}

public String getName() {
	return name;
}

public void setName(String name) {
	this.name = name;
}

public Date getDate() {
	return date;
}

public void setDate(Date date) {
	this.date = date;
}

public String getManagerName() {
	return ManagerName;
}

public void setManagerName(String managerName) {
	ManagerName = managerName;
}

@Override
public String toString() {
	return "CaPpmEmp [PsId=" + PsId + ", name=" + name + ", date=" + date + ", ManagerName=" + ManagerName + "]";
}

}
