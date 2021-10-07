package practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample {

	public static void main(String[] args) throws IOException {

		File f = new File("C:\\Users\\Hi\\eclipse-workspace\\Demo1\\xl\\vinitha.xlsx");

		FileInputStream s = new FileInputStream(f);

		Workbook sheet = new XSSFWorkbook(s);

		Sheet sheet2 = sheet.getSheet("Sheet1");

		for (int i = 0; i < sheet2.getPhysicalNumberOfRows(); i++) {

			Row row = sheet2.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

				Cell cell = row.getCell(j);

				int cellType = cell.getCellType();

				if (cellType == 1) {

					String string = cell.getStringCellValue();

					System.out.print(string + " ");
				}

				if (cellType == 0) {

					if (DateUtil.isCellDateFormatted(cell)) {
						Date date = cell.getDateCellValue();
						SimpleDateFormat d = new SimpleDateFormat("dd-MM-yyyy");
						String format = d.format(date);
						System.out.print(format);

					}

					double num = cell.getNumericCellValue();

					long l = (long) num;

					String value = String.valueOf(l);

					System.out.print(value + " ");

				}

			}
			System.out.println();
		}
	}

}
