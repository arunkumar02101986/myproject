package myExtractor.newCode;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewCode {
	private static String input_excelFile_path = "E:\\input\\input.xlsx";
	private static String input_zipFile_path = "E:\\input";
	private static String output_excelFile_path = "E:\\FinalFile.csv";
	static XSSFRow row;
	private static String header;

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream(new File(input_excelFile_path));
		PrintWriter writer = new PrintWriter(new File(output_excelFile_path));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = spreadsheet.iterator();
		int countforheader = 0;
		StringBuilder builder = null;
		File[] files = new File(input_zipFile_path).listFiles();

		while (rowIterator.hasNext()) {
			builder = new StringBuilder();
			row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					builder.append(Double.valueOf(cell.getNumericCellValue()).intValue() + ",");
					break;

				case Cell.CELL_TYPE_STRING:
					builder.append(cell.getStringCellValue() + ",");
					break;
				}
			}
			StringBuilder finalBuilder = new StringBuilder();
			if (countforheader == 0) {
				countforheader++;
				header = builder.toString() +"mFileNamesInTheZip" + "\n";
				writer.write(header);
			}
			String data = builder.toString();
			String[] splitData = data.split(",");
			try (FileInputStream fis1 = new FileInputStream(output_excelFile_path);
					BufferedInputStream bis = new BufferedInputStream(fis1);
					ZipInputStream zis = new ZipInputStream(bis)) {
				for (File file : files) {
					String fileName = file.getName();// .replaceAll("[^a-zA-Z0-9]", "").trim();
					if (fileName.contains(splitData[2]) && fileName.contains("zip")) {
						ZipFile zipFile = new ZipFile(file.getCanonicalPath());
						Enumeration zipEntries = zipFile.entries();
						String fname;
						while (zipEntries.hasMoreElements()) {
							fname = ((ZipEntry) zipEntries.nextElement()).getName();
							finalBuilder.append(builder.toString() + fname);
							writer.write(finalBuilder.toString()+"\n");
							System.out.println(finalBuilder.toString());
							finalBuilder.setLength(0);
						}
					}

				}
			} catch (Exception e) {
				e.printStackTrace();
			}

		}
		writer.close();
		fis.close();

	}

}
