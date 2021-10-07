package practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.SimpleFormatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SampleTwo {

	public static void main(String[] args) throws Throwable {

		File file = new File("C:\\Users\\Hi\\eclipse-workspace\\Demo1\\JaganInfo\\Jagan.xlsx");

		FileInputStream fileInputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(fileInputStream);

		Sheet sheet = workbook.getSheet("Sheet1");

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

				Cell cell = row.getCell(j);

				int cellType = cell.getCellType();

				if (cellType == 1) {

					String stringCellValue = cell.getStringCellValue();

					System.out.print(stringCellValue + " ");

				}

				if (cellType == 0) {

					if (DateUtil.isCellDateFormatted(cell)) {

						Date date = cell.getDateCellValue();

						SimpleDateFormat format = new SimpleDateFormat("dd-MM-yy");

						String da = format.format(date);

						System.out.print(da + " ");

					}

					double num = cell.getNumericCellValue();

					long l = (long) num;

					String v = String.valueOf(l);

					System.out.print(v + " ");

				}

			}
			System.out.println();
		}

	}

}
