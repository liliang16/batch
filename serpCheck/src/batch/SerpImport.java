package batch;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.security.GeneralSecurityException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.ResourceBundle;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * 导入serp excel表
 * 
 */
public class SerpImport {

	public static final String SERPDATA = "T_SERP_DATA";

	/** server address */
	static String url;
	/** MYSQL JDBC driver */
	static String driver;
	/** username */
	static String user;
	/** password */
	static String password;
	/** excelPassword */
	static String excelPassword = "";
	/** isFirstFileFlg */
	static String isFirstFileFlg;

	static String[] array_T_SERP_DATA_name = { "项目编号", "项目名称", "PM", "员工编号", "员工" };

	public static void main(String[] args) throws Exception {

		driver = args[2]; // MYSQL JDBC driver
		url = args[3]; // server address
		user = args[4]; // username
		password = args[5].toString(); // password
		isFirstFileFlg = args[6]; // isFirstFileFlg y:yes n:no
		if (args.length > 7) {
			excelPassword = args[7]; // excelPassword
		}

		System.out.println("server address:" + url);
		System.out.println("MYSQL JDBC driver:" + driver);
		System.out.println("username:" + user);
		System.out.println("password:" + password);
		// Excel file path
		String excelPath = args[0];
		// Table Name
		String tableName = args[1];
		System.out.println("Excel file path:" + args[0]);

		toMySQL(excelPath, tableName, excelPassword, isFirstFileFlg);
	}

	/*
	 * Implement the function of inserting Excel into MySQL database
	 */
	public static void toMySQL(String excelPath, String tableName, String excelPassword, String isFirstFileFlg)
			throws Exception {

		final ResourceBundle bundle = ResourceBundle.getBundle("excel");
		Integer[] array_T_SERP_DATA_int = StringToInteger(bundle.getString("array_T_SERP_DATA_int").split(","));

		Workbook wb = null;

		// Load the MYSQL JDBC driver program
		try {
			Class.forName(driver);
			System.out.println("The driver is loaded successfully!");
		} catch (Exception e) {
			System.out.print("The driver failed to load!");
			e.printStackTrace();
			throw new Exception(e);
		}

		try {
			// Database Connection
			Connection connect = DriverManager.getConnection(url, user, password);
			// Get an object that executes the sql statement based on the
			// connection
			Statement stmt = connect.createStatement();

			System.out.println("The database connection is successful!");

			File excel = new File(excelPath);
			if (excel.isFile() && excel.exists()) {
				// "//." is a special character that needs to be escaped
				String[] split = excel.getName().split("\\.");
				// Judge based on the file suffix (xls/xlsx)
				if ("" == excelPassword && "xls".equals(split[1])) {
					FileInputStream fis = new FileInputStream(excel);
					wb = new HSSFWorkbook(fis);
				} else if ("" == excelPassword && "xlsx".equals(split[1])) {
					wb = new XSSFWorkbook(excel);
				} else if ("" != excelPassword && "xls".equals(split[1])) {
					org.apache.poi.hssf.record.crypto.Biff8EncryptionKey.setCurrentUserPassword(excelPassword);
					FileInputStream fis = new FileInputStream(excel);
					wb = new HSSFWorkbook(fis);
					fis.close();
				} else if ("" != excelPassword && "xlsx".equals(split[1])) {
					InputStream inp = new FileInputStream(excelPath);
					POIFSFileSystem pfs = new POIFSFileSystem(inp);
					inp.close();
					EncryptionInfo encInfo = new EncryptionInfo(pfs);
					Decryptor decryptor = Decryptor.getInstance(encInfo);
					try {
						decryptor.verifyPassword(excelPassword);
						wb = new XSSFWorkbook(decryptor.getDataStream(pfs));
					} catch (GeneralSecurityException e) {
						e.printStackTrace();
						throw new Exception(e);
					}
				} else {
					System.out.println("The file is not excel!");
					return;
				}

				System.out.println("File reading begins");

				// read sheet 0
				Sheet sheetAt = wb.getSheetAt(0);
				// T_SERP_DATA table related SQL
				String sqlInsert_T_SERP_DATA = "replace into T_SERP_DATA (PROJECT_ID,PROJECT_NAME,PM_NAME,EMPLOYEE_NO,"
						+ "EMPLOYEE_NAME,COST_MANHOUR,PROJECT_YEAR,PROJECT_MONTH,ENT_DT,ENT_ID," + "DEL_FG) values (";

				SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				// Current system date
				String currentDate = dateFormat.format(new Date());
				// Current system month
				String nextMonth = "";

				SimpleDateFormat sdf = new SimpleDateFormat("MM");

				Calendar calendar = Calendar.getInstance();

				// get next month

				calendar.add(Calendar.MONTH, 1);

				nextMonth = sdf.format(calendar.getTime());

				System.out.println("next month：" + nextMonth);

				// userID
				String userID = "SERPImport";
				// Ent date and ID
				String sqlEntDateID = "'" + currentDate + "','" + userID + "',";
				// delete flg
				String delFlg = "0";
				String sqlNull = ");";
				String sqlOthers = sqlEntDateID + delFlg;

				List<Integer> list_T_SERP_DATA = new ArrayList<Integer>(Arrays.asList(array_T_SERP_DATA_int));

				List<Integer> list_int = null;
				// the column name
				int columnNameIndex = 0;
				// The first line is the column name, so don't read
				int firstRowIndex = 0;
				// Total number of rows
				int lastRowIndex = sheetAt.getLastRowNum();
				// Number of rows inserted
				int insertRowIndex = 0;
				// row Number of the fist year and month
				int firstCMHCellIndex = 10;
				// row Number of the fist COST_MANHOUR
				int cMHCellIndex = firstCMHCellIndex;
				// COST_MANHOUR for sql
				String sqlYM = "";
				String sqlYear = "";

				if (tableName.equals(SERPDATA)) {
					list_int = list_T_SERP_DATA;
					columnNameIndex = sheetAt.getFirstRowNum();
					firstRowIndex = sheetAt.getFirstRowNum() + 2;
				}

				Row rowColumnName = sheetAt.getRow(columnNameIndex);
				if (rowColumnName != null) {
					// get year and month of the fist cost man hour
					String firstMonthYear = rowColumnName.getCell(firstCMHCellIndex).toString();
					if (firstMonthYear.length() == 6) {
						String firstMonth = firstMonthYear.substring(4);
						String cellNextMonth = firstMonthYear;
						if (!nextMonth.equals(firstMonth)) {
							for (int i = 2; i < 12; i++) {
								cMHCellIndex = cMHCellIndex + 2;
								cellNextMonth = rowColumnName.getCell(cMHCellIndex).toString();
								if (cellNextMonth == null || cellNextMonth == "") {
									System.out.println("The year and month of data columns is null!");
									System.exit(1);
								}
								if (nextMonth.equals(cellNextMonth.substring(4))) {
									break;
								}
							}
						}
						sqlYM = "'" + cellNextMonth.substring(0, 4) + "','" + cellNextMonth.substring(4) + "',";
						sqlYear = cellNextMonth.substring(2, 4);
						// add the cell index of COST_MANHOUR
						list_int.add(cMHCellIndex);
						System.out.println("The list_int" + list_int);
					} else {
						System.out.println("The year and month of data columns is incorrect!" + firstMonthYear);
						System.exit(1);
					}
				}

				System.out.println("Traversal begins");
				String delSQL = "TRUNCATE T_SERP_DATA;";
				try {
					stmt.executeUpdate(delSQL);
				} catch (SQLException e) {
					System.out.println("TRUNCATE failed!");
					e.printStackTrace();
					extracted(e);
				}
				for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) { // Traversal
					
					boolean checkYear = false;
					// rows
					String sql = "";
					Row row = sheetAt.getRow(rIndex);
					if (row != null) {
						int firstCellIndex = row.getFirstCellNum();
						int lastCellIndex = row.getLastCellNum();
						for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) { // Traversal
							// column
							String sqlValue = "";

							Cell cell = row.getCell(cIndex);
							
							// Remove the blank column
							if (!list_int.contains(cIndex)) {
								continue;
							} else {
								if (cell != null && cell.toString() != "") {
									String cellValue = getCellValue(cell);
									//check 是否是跨年错误数据 —— 终止本次for循环
									if (cIndex == 5 && !sqlYear.equals(cellValue.substring(cellValue.length()-6, cellValue.length()-4))) {
										checkYear = true;
										break;
									}
									// projectID里的"O"改成"P"
									if (cIndex == 5 && "O".equals(
											cellValue.substring(cellValue.length() - 4, cellValue.length() - 3))) {
										String newCcellValue = cellValue.substring(0, cellValue.length() - 4) + "P"
												+ cellValue.substring(cellValue.length() - 3);
										sqlValue += "'" + newCcellValue + "'";
									} else {
										sqlValue += "'" + cellValue + "'";
									}
								} else {
									sqlValue += "NULL";
								}
								sql += sqlValue + ",";
							}
						}
						//check 是否是跨年错误数据 —— 跳过当前条件
						if(checkYear){
							continue;
						}
						// Insert the full SQL of the T_SERP_DATA table
						if (tableName.equals(SERPDATA)) {
							sql = sqlInsert_T_SERP_DATA + sql + sqlYM + sqlOthers + sqlNull;
						}

						System.out.println("Insert the " + rIndex + " data");

						System.out.println(sql);

						try {
							stmt.executeUpdate(sql);
						} catch (SQLException e) {
							System.out.println("Insert failed!");
							e.printStackTrace();
							extracted(e);
						}
						insertRowIndex++;
					}
				}

				lastRowIndex = lastRowIndex - 1;

				if (insertRowIndex == lastRowIndex) {
					System.out.println("All data was inserted successfully!");
				} else {
					System.out.println("Some or all of the data insertion failed!");
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception(e);
		} finally {
			if (null != wb) {
				try {
					wb.close();
				} catch (IOException e) {
					System.out.println("The flow failure failed!");
					e.printStackTrace();
					throw new Exception(e);
				}
			}
			System.out.println("End of all processing");
			// Program termination
			System.exit(0);
		}
	}

	private static void extracted(SQLException e) throws Exception {
		throw new Exception(e);
	}

	/*
	 * Cell content acquisition
	 */
	public static String getCellValue(Cell cell) throws UnsupportedEncodingException {
		String value = "";
		switch (cell.getCellType()) {
		case STRING:
			String cellValue = cell.getRichStringCellValue().getString();
			value = new String(cellValue.getBytes(), "utf-8");
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				value = new SimpleDateFormat("yyyy/MM/dd").format(cell.getDateCellValue());
			} else {
				value = String.valueOf(Double.valueOf(cell.getNumericCellValue()).longValue());
			}
			break;
		case BOOLEAN:
			value = String.valueOf(cell.getBooleanCellValue());
			break;
		case FORMULA:
			value = cell.getCellFormula();
			break;
		case BLANK:
			break;
		default:
		}
		if (value == null)
			return null;
		return value.replaceAll("[　*| *| *|//s*]*", "");
	}

	public static Integer[] StringToInteger(String[] arrs) {
		Integer[] ints = new Integer[arrs.length];
		for (int i = 0; i < arrs.length; i++) {
			ints[i] = Integer.parseInt(arrs[i]);
		}
		return ints;
	}
}