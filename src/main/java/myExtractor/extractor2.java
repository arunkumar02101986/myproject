package myExtractor;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class getAllfilesFromExcel {
	private final static String inputFilePath = "E:\\input\\input.xlsx";
	private final static String outputFilePath = "E:\\input\\output.csv";
	static XSSFRow row;
	private static String header;

	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream(new File(inputFilePath));
		PrintWriter writer = new PrintWriter(new File(outputFilePath));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = spreadsheet.iterator();
		int countforheader = 0;
		StringBuilder builder = null;

		while (rowIterator.hasNext()) {
			row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			builder = new StringBuilder();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					builder.append(Double.valueOf(cell.getNumericCellValue()).intValue() + ",");
					break;

				case Cell.CELL_TYPE_STRING:
					builder.append(cell.getStringCellValue()+ ",");
					break;
				}
			}
			if (countforheader == 0) {
				countforheader++;
				header = builder.toString();
				System.out.println(header);
				writer.write(header + "FileName \n");
				continue;
			} else {
				String input1 = builder.toString();
				String secondValue = input1.split(",")[1];
				String fileName = "E:\\input\\" + secondValue + ".zip";
				String finalString = null;
				try (FileInputStream fis1 = new FileInputStream(fileName);
						BufferedInputStream bis = new BufferedInputStream(fis1);
						ZipInputStream zis = new ZipInputStream(bis)) {
					finalString = null;
					ZipEntry ze;
					while ((ze = zis.getNextEntry()) != null) {
//						System.out.println(input1 + "," + ze.getName());
						finalString = input1 + ze.getName();
						System.out.println(finalString);
						writer.write(finalString+"\n");
					}
					
				} catch (Exception e) {
					// TODO: handle exception
//					System.out.println(input1 + ",File not found" + e.getMessage());
					finalString = input1 + "File not found" + e.getMessage();
					System.out.println(finalString);
					writer.write(finalString+"\n");
				}
//				input1 = input1 + "\n";
//				System.out.println(input1);

			}
		}
		writer.close();
		fis.close();
	}
}
