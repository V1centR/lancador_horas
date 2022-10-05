package main.java.com.lancadorhoras.utils;

import java.util.ArrayList;
import java.util.List;

import main.java.com.lancadorhoras.model.ColumnsSheet;

public class SheetManagement {
	
	public List<ColumnsSheet> generateDataLines(int year, int month){
		
		DateProcess dateProcess = new DateProcess();
		
		List<ColumnsSheet> lancamentos = new ArrayList<ColumnsSheet>();
		
		for(int i = 1; i <= dateProcess.monthDays(year,month);i++) {
			lancamentos.add(new ColumnsSheet(String.format("%02d", i) + "/" + String.format("%02d", month)  + "/" + year, "08:" + dateProcess.randomMinute(20),"12:00","13:" + dateProcess.randomMinute(5),"17:00","0","0","08:00","-"));
		}
		
		return lancamentos;
	}
	
	public String[] setupHeaders() {
		
		return new String[] { "Data", "Entrada", "Saida","Entrada","Saida","ExtraEntrada","ExtraSaida","TotalHoras","Obs" };
	}

}
