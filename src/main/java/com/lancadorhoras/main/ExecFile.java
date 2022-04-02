package main.java.com.lancadorhoras.main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import main.java.com.lancadorhoras.model.ColumnsSheet;
import main.java.com.lancadorhoras.utils.DateProcess;

public class ExecFile {
	
	public static void main(String[] args) {
		
		
		//private final Path pathTmp = Paths.get("output");
		
		
		int year = 2022;
		int month = 04;
		
		String currentMonth = month + "-" + year;
		String fileName = "./output/planilha_horas_" + currentMonth+".xls";
		
		//XSSFWorkbook    workbook    = new XSSFWorkbook();
		HSSFWorkbook workbook = new HSSFWorkbook();
		//XSSFSheet sheetLancamentos = workbook.createSheet(currentMonth);
		HSSFSheet sheetLancamentos = workbook.createSheet(currentMonth);
        
        
        
        
        DateProcess dateProcess = new DateProcess();
        
       // System.out.println("total dias mês::: " + dateProcess.monthDays());
        
        
        List<ColumnsSheet> lancamentos = new ArrayList<ColumnsSheet>();
        
        for(int i = 1; i< dateProcess.monthDays(year,month);i++) {
        	lancamentos.add(new ColumnsSheet(String.format("%02d", i) + "/04/2022", "09:00","12:00","13:30","18:00","0","0","08:00","-"));
        }
		
        //SetupHeader
        String[] headers = new String[] { "Data", "Entrada", "Saida","Entrada","Saida","ExtraEntrada","ExtraSaida","TotalHoras","Obs" };
        
        Row firstrow = sheetLancamentos.createRow(0);
        int cellHeadernum = 0;
        for (Object obj : headers) {
			Cell cell = firstrow.createCell(cellHeadernum++);
			if (obj instanceof String)
				cell.setCellValue((String) obj);
			else if (obj instanceof Integer)
				cell.setCellValue((Integer) obj);
		}

        //SetupColor
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.PALE_BLUE.getIndex());
        style.setFillPattern(style.getFillPattern().ALT_BARS);    	
        
        
        int rownum = 1;
        for (ColumnsSheet lineData : lancamentos) {
        	int cellnum = 0;
        	
        	Row row = sheetLancamentos.createRow(rownum++);
        	
        	
        	//Dia
        	Cell cellDia = row.createCell(cellnum++);
        	cellDia.setCellValue(lineData.getDia());
        	cellDia.setCellStyle(style);
        	
        	//Entrada
        	Cell cellEntrada = row.createCell(cellnum++);
        	cellEntrada.setCellValue(lineData.getEntrada());
        	
        	//Saida
        	Cell cellSaida = row.createCell(cellnum++);
        	cellSaida.setCellValue(lineData.getSaida());
        	
        	//Entrada2
        	Cell cellEntrada2 = row.createCell(cellnum++);
        	cellEntrada2.setCellValue(lineData.getEntrada2());
        	
        	//Saida2
        	Cell cellSaida2 = row.createCell(cellnum++);
        	cellSaida2.setCellValue(lineData.getSaida2());
        	
        	//ExtraEntrada
        	Cell cellExtraEntrada = row.createCell(cellnum++);
        	cellExtraEntrada.setCellValue(lineData.getExtraEntrada());
        	
        	//ExtraSaida
        	Cell cellExtraSaida = row.createCell(cellnum++);
        	cellExtraSaida.setCellValue(lineData.getExtraSaida());
        	
        	//TotalHoras
        	Cell cellTotalHoras = row.createCell(cellnum++);
        	cellTotalHoras.setCellValue(lineData.getTotalHoras());
        	
        	//Obs
        	Cell cellObs = row.createCell(cellnum++);
        	cellObs.setCellValue(lineData.getObs());
        	
        }
        
        try {
            FileOutputStream out = new FileOutputStream(new File(fileName));
            workbook.write(out);
            out.close();
            System.out.println("Arquivo Excel criado com sucesso!");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
               System.out.println("Arquivo não encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
               System.out.println("Erro na edição do arquivo!");
        }

	}

}
