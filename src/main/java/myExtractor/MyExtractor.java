package myExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.lingala.zip4j.ZipFile;

public class MyExtractor {
	public static void main(String[] args) throws IOException {
		Map<String, String> myMap = new HashMap<>();
		Files.lines(Paths.get("myprop.properties")).forEach(line -> {
			String[] splitlines = line.split("=");
			myMap.put(splitlines[0].trim(), splitlines[1].trim());
		});

		System.out.println(myMap);
		FileInputStream inputStream = new FileInputStream(new File(myMap.get("input-excel-file-path")));
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet datatypeSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = datatypeSheet.iterator();
		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			while (cellIterator.hasNext()) {
				
				Cell cell = cellIterator.next();
				if(cell.getColumnIndex()==Integer.parseInt(myMap.get("sheet-column-index"))) {
				if (cell.getRowIndex() == 0) continue;
				
				System.out.println(cell.getRowIndex() + " " + cell.getStringCellValue());

				List<File> filesListFromPath = Files.list(Paths.get(myMap.get("input-path"))).map(Path::toFile)
						.collect(Collectors.toList());
				filesListFromPath.forEach(file -> {
					System.out.println(" file " + file);
					if(file.getName().equalsIgnoreCase(cell.getStringCellValue()+"."+myMap.get("file-format"))){
						
						try {
							unzipFolderZip4j(file,myMap.get("output-path"),myMap);
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}
				});
			}
		}
		}
		inputStream.close();

	}
	  // it takes `File` as arguments
	  public static void unzipFolderZip4j(File file, String output,Map<String,String> myMap)
	        throws IOException {
		  if(!file.getName().contains(myMap.get("file-not-include"))) {
			  
		  }
	        

	  }
}
