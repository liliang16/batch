/**
 * 
 */
package batch;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author P0000316 根据导入的serp数据，检索项目管理里是否有没有的项目（PROJECT），没有的追加
 */
public class ProjectCheck {

	/**
	 * @param args
	 */

	/** server address */
	public static String dbUrl;
	/** MYSQL JDBC driver */
	private static String driver;
	/** DB user name */
	private static String dbUser;
	/** DB password */
	private static String dbPassword;

	public static void main(String[] args) throws Exception {

		driver = args[0]; // MYSQL JDBC driver
		dbUrl = args[1]; // server address
		dbUser = args[2]; // db user name
		dbPassword = args[3].toString(); // db password

		System.out.println("DB server address:" + dbUrl);
		System.out.println("MYSQL JDBC driver:" + driver);
		System.out.println("DB username:" + dbUser);
		System.out.println("DB password:" + dbPassword);

		Connection connect = null;
		Statement stmtl = null;
		Statement insStmtl = null;
		ResultSet rsTEW = null;

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
			connect = DriverManager.getConnection(dbUrl, dbUser, dbPassword);
			System.out.println("The database connection is successful!");

			// Base on SERP data, find out the missed Project
			String selNoProjSQL = "SELECT serp.PROJECT_ID PROJECT_ID, serp.PROJECT_YEAR YEAR, serp.PM_NAME PM_NAME,serp.PROJECT_NAME PROJECT_NAME,employee.EMPLOYEE_NO PM_EMPLOYEE_NO"
					+ " FROM (SELECT PROJECT_ID, PROJECT_YEAR, PM_NAME, PROJECT_NAME FROM T_SERP_DATA WHERE DEL_FG = '0' GROUP BY PROJECT_ID,PROJECT_NAME,PM_NAME,PROJECT_YEAR,PROJECT_MONTH) serp"
					+ " LEFT JOIN M_PROJECT proj ON serp.PROJECT_ID = proj.PROJECT_ID AND serp.PROJECT_YEAR = proj.YEAR AND proj.DEL_FG = '0'"
					+ " LEFT JOIN M_EMPLOYEE employee ON serp.PM_NAME = employee.NAME AND employee.DEL_FG = '0'"
					+ " WHERE proj.PROJECT_ID IS NULL OR proj.YEAR IS NULL ORDER BY serp.PROJECT_ID,serp.PROJECT_NAME";

			System.out.println("select SQL: " + selNoProjSQL);

			// Run the select
			stmtl = connect.createStatement();
			rsTEW = stmtl.executeQuery(selNoProjSQL);
			insStmtl = connect.createStatement();

			String sql = "";
			String insertSQL = "replace into M_PROJECT (PM_EMPLOYEE_NO,YEAR,PROJECT_ID,PROJECT_NAME,CASE_NAME,ENT_DT,ENT_ID) values(";
			String strvalue = "";

			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			// Current system date
			String currentDate = dateFormat.format(new Date());
			String entID = "ProjUp";
			int errorcount = 0;

			while (rsTEW.next()) {
				strvalue = "";
				strvalue = strvalue + "'" + rsTEW.getString("PM_EMPLOYEE_NO") + "',";
				strvalue = strvalue + "'" + rsTEW.getString("YEAR") + "',";
				strvalue = strvalue + "'" + rsTEW.getString("PROJECT_ID") + "',";
				strvalue = strvalue + "'" + rsTEW.getString("PROJECT_NAME") + "',";
				strvalue = strvalue + "'0',";
				strvalue = strvalue + "'" + currentDate + "','" + entID + "')";
				sql = insertSQL + strvalue;
				System.out.println("Inserting into M_PROJECT-project_id:" + rsTEW.getString("PROJECT_ID") + ", PM_No:"
						+ rsTEW.getString("PM_EMPLOYEE_NO"));
				System.out.println("insert SQL:" + sql);

				try {
					insStmtl.executeUpdate(sql);
				} catch (SQLException e) {
					System.out.println("Insert failed!");
					errorcount = errorcount + 1;
					e.printStackTrace();
					throw new Exception(e);
				}
			}
			if (errorcount == 0)
				System.out.println("All New projects inserted successfully");
			else
				System.out.println("some or all of new projects insert failed");

		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception(e);
		} finally {
			if (rsTEW != null) {
				try {
					rsTEW.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
				rsTEW = null;

			}
			if (stmtl != null) {
				try {
					stmtl.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			if (insStmtl != null) {
				try {
					stmtl.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			if (connect != null) {
				try {
					connect.close();
				} catch (Exception e) {
					e.printStackTrace();
				}

			}
		}

	}

}
