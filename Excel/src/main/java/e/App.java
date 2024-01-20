package e;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.sql.Types;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	public static void main(String[] args) {
		System.out.println("Hello World!");

		String url = "jdbc:postgresql://localhost:5432/postgres";
		String username = "postgres";
		String password = "root";
		File file = new File("G:\\file_example_XLSX_50.xlsx");

		try {
			FileInputStream excelFile = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(excelFile);

			Connection connection = DriverManager.getConnection(url, username, password);
			Statement statement = connection.createStatement();

			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFRow headerRow = sheet.getRow(0);

			StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS excelDemo (");
			for (Cell cell : headerRow) {
				String columnName = cell.getStringCellValue();
				String dataType = dataType(sheet.getRow(1).getCell(cell.getColumnIndex()));
				columnName = columnName.replaceAll("\\s", "_");
				columnName = columnName.replaceAll("[():/]", "_");
				columnName = columnName.replaceAll("[.#\\-]", "");
				columnName = columnName.replaceAll("1st", "first");
				createTableQuery.append(columnName).append(" ").append(dataType).append(", ");
			}
			createTableQuery.delete(createTableQuery.length() - 2, createTableQuery.length());// remove last comma and
																								// space
			createTableQuery.append(");");
			statement.executeUpdate(createTableQuery.toString());
			System.out.println(createTableQuery);
			System.out.println("Table Created Successfully!");

			int numberofRows = sheet.getPhysicalNumberOfRows();
			for (int i = 1; i < numberofRows; i++) {
				Row row = sheet.getRow(i);
				StringBuilder insertQuery = new StringBuilder("INSERT INTO excelDemo VALUES(");

				for (int j = 0; j < headerRow.getPhysicalNumberOfCells(); j++) {
					Cell cell = row.getCell(j);
					insertQuery.append("?,");
				}
				insertQuery.delete(insertQuery.length() - 1, insertQuery.length());
				insertQuery.append(");");

				try {
					PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString());
					for (int j = 0; j < headerRow.getPhysicalNumberOfCells(); j++) {
						Cell datacell = row.getCell(j);

						if (datacell != null) {
							switch (datacell.getCellType()) {
							case STRING:
								preparedStatement.setString(j + 1, datacell.getStringCellValue());
								break;
							case NUMERIC:
								if (DateUtil.isCellDateFormatted(datacell)) {
									java.util.Date dataCellValue = datacell.getDateCellValue();
									preparedStatement.setDate(j + 1, new java.sql.Date(dataCellValue.getTime()));
								} else {
									preparedStatement.setInt(j + 1, (int) datacell.getNumericCellValue());
								}
								break;
							case BOOLEAN:
								preparedStatement.setBoolean(j + 1, datacell.getBooleanCellValue());
								break;
							default:
								preparedStatement.setNull(j + 1, Types.NULL);
							}
						} else {
							preparedStatement.setNull(j + 1, Types.NULL);
						}
					}

					preparedStatement.executeUpdate();
					System.out.println(insertQuery);
					System.out.println("Data inserted Successfully!");
				} catch (Exception e) {

				}
			}
		} catch (Exception e) {
			System.out.println("Error: " + e.getMessage());
			e.printStackTrace();
		}
	}

	private static String dataType(Cell cell) {
		if (cell.getCellType() == CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				return "DATE";
			} else {
				return "INT";
			}
		} else if (cell.getCellType() == CellType.STRING) {
			return "VARCHAR(255)";
		} else if (cell.getCellType() == CellType.BOOLEAN) {
			return "BOOLEAN";
		} else {
			return "VARCHAR(255)";
		}
	}
}
