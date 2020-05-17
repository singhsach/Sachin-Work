package com.java.Email;

import java.time.LocalDate;
import java.time.Month;

public class TimesheetMonth {

	public Month getCurrentMonth() {

		
		LocalDate currentdate = LocalDate.now().plusMonths(-1);
		Month currentMonth = currentdate.getMonth();
		System.out.println("Current month: " + currentMonth);

		return currentMonth;
	}
}
