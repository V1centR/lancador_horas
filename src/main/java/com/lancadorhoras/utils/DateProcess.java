package main.java.com.lancadorhoras.utils;

import java.time.LocalDate;
import java.util.Random;

public class DateProcess {

	public int monthDays(int year, int month) {

		LocalDate date = LocalDate.of(year, month, 1);
		return date.lengthOfMonth();
	}

	public String randomMinute(int maxNumber) {

		Random randNumber = new Random();

		int number = 1 + randNumber.nextInt(maxNumber);

		return String.format("%02d", number);
	}

}
