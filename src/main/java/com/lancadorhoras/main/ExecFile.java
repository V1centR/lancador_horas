package main.java.com.lancadorhoras.main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import main.java.com.lancadorhoras.model.ColumnsSheet;
import main.java.com.lancadorhoras.utils.SheetManagement;

public class ExecFile {

	public static void main(String[] args) {

		// private final Path pathTmp = Paths.get("output");

		//int year = 2022;
		int year = Integer.valueOf(args[0]);

		//String currentMonth = month + "-" + year;
		String fileName = "./output/planilha_horas_" + year + ".xls";
		HSSFWorkbook workbook = new HSSFWorkbook();

		SheetManagement sm = new SheetManagement();

		// because one year have 12 months
		// this "for" setup all current sheet data
		for (int i = 1; i <= 12; i++) {

			// Create Sheet
			HSSFSheet sheetLancamentos = workbook.createSheet(String.format("%02d", i) + "-" + year);
			// var lancamentos
			List<ColumnsSheet> lancamentos = sm.generateDataLines(year, i);

			Row firstrow = sheetLancamentos.createRow(0);
			int cellHeadernum = 0;
			for (Object obj : sm.setupHeaders()) {
				Cell cell = firstrow.createCell(cellHeadernum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}

			// SetupColor
			HSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.PALE_BLUE.getIndex());
			style.setFillPattern(style.getFillPattern().ALT_BARS);

			int rownum = 1;
			for (ColumnsSheet lineData : lancamentos) {
				int cellnum = 0;

				Row row = sheetLancamentos.createRow(rownum++);

				// Dia
				Cell cellDia = row.createCell(cellnum++);
				cellDia.setCellValue(lineData.getDia());
				cellDia.setCellStyle(style);

				// Entrada
				Cell cellEntrada = row.createCell(cellnum++);
				cellEntrada.setCellValue(lineData.getEntrada());

				// Saida
				Cell cellSaida = row.createCell(cellnum++);
				cellSaida.setCellValue(lineData.getSaida());

				// Entrada2
				Cell cellEntrada2 = row.createCell(cellnum++);
				cellEntrada2.setCellValue(lineData.getEntrada2());

				// Saida2
				Cell cellSaida2 = row.createCell(cellnum++);
				cellSaida2.setCellValue(lineData.getSaida2());

				// ExtraEntrada
				Cell cellExtraEntrada = row.createCell(cellnum++);
				cellExtraEntrada.setCellValue(lineData.getExtraEntrada());

				// ExtraSaida
				Cell cellExtraSaida = row.createCell(cellnum++);
				cellExtraSaida.setCellValue(lineData.getExtraSaida());

				// TotalHoras
				Cell cellTotalHoras = row.createCell(cellnum++);
				cellTotalHoras.setCellValue(lineData.getTotalHoras());

				// Obs
				Cell cellObs = row.createCell(cellnum++);
				cellObs.setCellValue(lineData.getObs());

			}

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
