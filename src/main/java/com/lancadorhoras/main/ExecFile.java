package main.java.com.lancadorhoras.main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import main.java.com.lancadorhoras.model.ColumnsSheet;

public class ExecFile {
	
	public static void main(String[] args) {
		
		
		//private final Path pathTmp = Paths.get("output");
		
		String currentMonth = "Abril";
		String fileName = "./output/planilha_horas_" + currentMonth+".xls";
		
		HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheetLancamentos = workbook.createSheet(currentMonth);
        
        
        List<ColumnsSheet> lancamentos = new ArrayList<ColumnsSheet>();
        
        //laço aqui gerando os valores e populando o array
        lancamentos.add(new ColumnsSheet("01/04/2022", "09:00","12:00","13:30","18:00","0","0","08:00","-"));
        lancamentos.add(new ColumnsSheet("02/04/2022", "09:00","12:00","13:30","18:00","0","0","08:00","-"));
        lancamentos.add(new ColumnsSheet("03/04/2022", "09:00","12:00","13:30","18:00","0","0","08:00","-"));
		
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
        
        
        int rownum = 1;
        for (ColumnsSheet lineData : lancamentos) {
        	int cellnum = 0;
        	
        	Row row = sheetLancamentos.createRow(rownum++);
        	
        	//Dia
        	Cell cellDia = row.createCell(cellnum++);
        	cellDia.setCellValue(lineData.getDia());
        	
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
