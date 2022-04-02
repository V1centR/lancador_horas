package main.java.com.lancadorhoras.utils;

import java.time.YearMonth;
import java.util.Calendar;

public class DateProcess {
	
	public int monthDays(int year,int month) {
		
		Calendar calendar = Calendar.getInstance();
		
		calendar.set(Calendar.YEAR, year);
		calendar.set(Calendar.MONTH, month);
		//int numDays = calendar.getActualMaximum(Calendar.DATE);
		
		return calendar.getActualMaximum(Calendar.DATE);
		
	}
	
	public String weekend() {
		
		
		
		return null;
	}

}
