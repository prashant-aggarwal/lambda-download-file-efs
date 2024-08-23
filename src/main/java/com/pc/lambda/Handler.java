package com.pc.lambda;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.amazonaws.services.lambda.runtime.Context;
import com.amazonaws.services.lambda.runtime.LambdaLogger;
import com.amazonaws.services.lambda.runtime.RequestHandler;

public class Handler implements RequestHandler<String, String> {

	private static LambdaLogger logger = null;
	private Map<String, Map<Integer, List<String>>> mMapSheetToDataMap = new TreeMap<>();

	public String handleRequest(String input, Context context) {

		logger = context.getLogger();

		// Read all of the environment variables
		Boolean readFile = Boolean.valueOf(System.getenv("readFile"));
		Boolean writeFile = Boolean.valueOf(System.getenv("writeFile"));
		Boolean writeReadLogs = Boolean.valueOf(System.getenv("writeReadLogs"));
		String fileDirectory = System.getenv("efsFolder");
		String fileName = System.getenv("fileName");
		String newFileName = System.getenv("newFileName");
		Boolean copyFile = Boolean.valueOf(System.getenv("copyFile"));

		// Prepare the file name to be processed
		String fullFilePath = fileDirectory + fileName;

		logger.log("\nreadFile => " + readFile);
		logger.log("\nwriteFile => " + writeFile);
		logger.log("\nwriteReadLogs => " + writeReadLogs);
		logger.log("\nfileDirectory => " + fileDirectory);
		logger.log("\nfileName => " + fileName);
		logger.log("\nnewFileName => " + newFileName);
		logger.log("\nfullFilePath => " + fullFilePath);
		logger.log("\ncopyFile => " + copyFile);

		// Create the file directory if it doesn't exist
		Path dir = Paths.get(fileDirectory);
		if (Files.exists(dir)) {
			logger.log("\n!! Directory Available !!");
		} else {
			logger.log("\n!! Directory Unavailable !!");
			try {
				logger.log("\n!! Creating Directory !!");
				Files.createDirectories(dir);
				logger.log("\n!! Created Directory !!");
			} catch (IOException e) {
				logger.log("\n!! Error while creating directory !!");
				return "\\n!! Error while creating directory !!";
			}
		}

		// Verify whether the file to be processed exists or not
		Path newFilePath = Paths.get(fullFilePath);
		Boolean fileExists = Files.exists(newFilePath);
		if (fileExists) {
			logger.log("\nFile is available: " + newFilePath.toAbsolutePath().toString());
		} else {
			logger.log("\nFile is unavailable: " + newFilePath.toAbsolutePath().toString());
		}

		try {
			if (copyFile) {
				readExcelFile(fullFilePath, writeReadLogs);

				fullFilePath = fileDirectory + newFileName;
				writeExcelFile(fullFilePath, copyFile);

				logger.log("\nFile copied successfully: " + fullFilePath);
				return "";
			}

			// Write File
			if (writeFile) {
				writeExcelFile(fullFilePath, copyFile);
				fileExists = true;
				logger.log("\nFile written successfully: " + fullFilePath);
			}

			// Read File
			if (fileExists && readFile) {
				readExcelFile(fullFilePath, writeReadLogs);
				logger.log("\nFile read successfully: " + fullFilePath);
			}
		} catch (Exception e) {
			logger.log("\nException while performing file operations: " + e.getMessage() + " \nStack Trace: "
					+ e.getStackTrace());
		}

		logger.log("\n\n");
		return "Processed successfully";
	}

	private void writeExcelFile(String filePath, Boolean copyFile) throws IOException {

		logger.log("\nInside writeExcelFile method.");

		// Blank workbook
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {

			logger.log("\nInitialized an instance of XSSFWorkbook");

			if (copyFile) {
				for (Map.Entry<String, Map<Integer, List<String>>> mapEntry : mMapSheetToDataMap.entrySet()) {
					String sheetName = mapEntry.getKey();
					XSSFSheet sheet = workbook.createSheet(sheetName);
					logger.log("\nCreated sheet: " + sheetName);

					Map<Integer, List<String>> mapSheetValues = mapEntry.getValue();
					// Iterate over data and write to sheet
					Set<Integer> keyset = mapSheetValues.keySet();
					int rownum = 0;
					for (Integer key : keyset) {
						Row row = sheet.createRow(rownum++);
						List<String> cellVaules = mapSheetValues.get(key);
						int cellnum = 0;
						for (String value : cellVaules) {
							Cell cell = row.createCell(cellnum++);
							cell.setCellValue(value);
						}
					}
				}
			} else {
				// Create a blank sheet
				XSSFSheet sheet = workbook.createSheet("Sample Data");

				logger.log("\nCreated blank sheet.");

				// This data needs to be written (Object[])
				Map<String, Object[]> data = new TreeMap<String, Object[]>();
				data.put("1", new Object[] { "ID", "NAME", "LASTNAME" });
				data.put("2", new Object[] { 1, "Amit", "Shukla" });
				data.put("3", new Object[] { 2, "Lokesh", "Gupta" });
				data.put("4", new Object[] { 3, "John", "Adwards" });
				data.put("5", new Object[] { 4, "Brian", "Schultz" });

				// Iterate over data and write to sheet
				Set<String> keyset = data.keySet();
				int rownum = 0;
				for (String key : keyset) {
					Row row = sheet.createRow(rownum++);
					Object[] objArr = data.get(key);
					int cellnum = 0;
					for (Object obj : objArr) {
						Cell cell = row.createCell(cellnum++);
						if (obj instanceof Integer)
							cell.setCellValue((Integer) obj);
						else
							cell.setCellValue(obj.toString());
					}
				}
			}

			logger.log("\nPopulated data.");

			// Write the workbook in file system
			try (FileOutputStream out = new FileOutputStream(new File(filePath))) {
				logger.log("\nInitialized an instance of FileOutputStream.");
				workbook.write(out);
			}

			logger.log("\nwriteExcelFile method executed.");
		}
	}

	private void readExcelFile(String filePath, Boolean writeReadLogs) throws FileNotFoundException, IOException {
		try (FileInputStream file = new FileInputStream(new File(filePath))) {
			logger.log("\nInside readExcelFile method.");

			// Create Workbook instance holding reference to .xlsx file
			try (XSSFWorkbook workbook = new XSSFWorkbook(file)) {
				workbook.forEach((sheet) -> {
					logger.log("\nSheet Name: " + sheet.getSheetName());
					logger.log("\nRow Count: " + sheet.getPhysicalNumberOfRows());
					
					Map<Integer, List<String>> sheetDataMap = new TreeMap<>();

					// Iterate through each rows one by one
					Iterator<Row> rowIterator = sheet.iterator();
					while (rowIterator.hasNext()) {
						Row row = rowIterator.next();
						List<String> listCellValues = new ArrayList<String>();

						// For each row, iterate through all the columns
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							listCellValues.add(cell.getStringCellValue());
							// Check the cell type and format accordingly
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_NUMERIC:
								if (writeReadLogs) {
									logger.log(cell.getNumericCellValue() + "\t");
								}
								break;
							default:
								if (writeReadLogs) {
									logger.log(cell.getStringCellValue() + "\t");
								}
								break;
							}
						}
						sheetDataMap.put(row.getRowNum(), listCellValues);
					}

					mMapSheetToDataMap.put(sheet.getSheetName(), sheetDataMap);
				});
			}

			logger.log("\nreadExcelFile method executed.");
		}
	}
}
