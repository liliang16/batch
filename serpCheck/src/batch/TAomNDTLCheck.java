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

import org.apache.commons.codec.binary.StringUtils;

/**
 * @author P0000316 根据导入的serp数据，检索受注管理里是否有没有的项目（PROJECT），没有的追加
 */
public class TAomNDTLCheck {

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
		ResultSet rsTEW2 = null;

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

			// 对比SERP和受注表，检索受注表中没有的项目（根据project_id和year）
			String selNoProjSQL = "SELECT DISTINCT IF(SUBSTR(serp.PROJECT_ID, 1, 1)='J','JP','DL') PRE_PRO_ID,SUBSTR(serp.PROJECT_ID, 3) PROJECT_NO,serp.PROJECT_YEAR PROJECT_YEAR,"
					+ "serp.PM_NAME PM_NAME,serp.PROJECT_NAME PROJECT_NAME,employee.EMPLOYEE_NO PM_NO FROM `T_SERP_DATA` serp"
					+ " LEFT JOIN `T_AOM_NDTL` aom ON SUBSTR(serp.PROJECT_ID,3) = SUBSTR(aom.PROJECT_ID,3) AND serp.PROJECT_YEAR=aom.`YEAR` AND serp.DEL_FG='0' AND aom.DEL_FG='0'"
					+ " LEFT JOIN M_EMPLOYEE employee ON serp.PM_NAME = employee.NAME AND employee.DEL_FG ='0'"
					+ " WHERE SUBSTR(aom.PROJECT_ID,3) IS NULL OR aom.`YEAR` IS NULL ORDER BY PROJECT_NO, PRE_PRO_ID";

			System.out.println("select SQL: " + selNoProjSQL);

			// Run the select
			stmtl = connect.createStatement();
			rsTEW = stmtl.executeQuery(selNoProjSQL);
			insStmtl = connect.createStatement();

			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			// Current system date
			String currentDate = dateFormat.format(new Date());
			String entID = "TAOMUp";
			int errorcount = 0;

			// 把检索结果的项目追加到受注表里
			while (rsTEW.next()) {
				String insertSQL = "replace into T_AOM_NDTL(PROJECT_ID,TEAM,PM_NO,WORK_PLACE,`YEAR`,ENT_DT,ENT_ID) values(";

				String idType = rsTEW.getString("PRE_PRO_ID");
				// 做成Insert语句中的PROJECT_ID的值
				String strProID1, strProID2, strProID3;
				if (StringUtils.equals(idType, "JP")) {
					// 日本项目 （J1XXXXXX,JPXXXXXX）
					strProID1 = "'J1" + rsTEW.getString("PROJECT_NO") + "',";
					strProID2 = "'JP" + rsTEW.getString("PROJECT_NO") + "',";
					strProID3 = "'JP" + rsTEW.getString("PROJECT_NO") + "',";

				} else {
					// 国内项目 （DLXXXXXX）
					strProID1 = "'DL" + rsTEW.getString("PROJECT_NO") + "',";
					strProID2 = "'DL" + rsTEW.getString("PROJECT_NO") + "',";
					strProID3 = "'DL" + rsTEW.getString("PROJECT_NO") + "',";
				}

				String strVPart1 = "";
				String strVPart2 = "";

				strVPart1 = strVPart1 + "'" + rsTEW.getString("PROJECT_NAME") + "',"; // TEAM
				strVPart1 = strVPart1 + "'" + rsTEW.getString("PM_NO") + "',"; // PM_NO
				strVPart2 = strVPart2 + "'" + rsTEW.getString("PROJECT_YEAR") + "',"; // YEAR
				strVPart2 = strVPart2 + "'" + currentDate + "','" + entID + "')"; // ENT_DT,ENT_ID

				String sql1 = insertSQL + strProID1 + strVPart1 + "'01'," + strVPart2; // JAPAN
				String sql2 = insertSQL + strProID2 + strVPart1 + "'02'," + strVPart2; // DALIAN
				String sql3 = insertSQL + strProID3 + strVPart1 + "'03'," + strVPart2; // WEIFANG

				// 第二种追加方式：
				// 检索受注管理里去年是否有该项目（ProjectID中的“年”有变化），如果有的话，根据去年的做成新的受注数据
				// 根据projectNo的倒数5,6位取得对应年，对应年-1生成前一年的projectNo
				StringBuilder strBProNo = new StringBuilder("");
				strBProNo.append(rsTEW.getString("PROJECT_NO"));
				String stroldY = Integer.toString(
						Integer.parseInt(strBProNo.substring(strBProNo.length() - 6, strBProNo.length() - 4)) - 1);
				String strOldProNo = strBProNo.replace(strBProNo.length() - 6, strBProNo.length() - 4, stroldY)
						.toString();

				// 检索前一年是否有这个project
				String sqlselOld;
				sqlselOld = "SELECT TEAM,BUSINESS_FIELD,CONTRACTUAL_RELATION,FIELD FROM `T_AOM_NDTL`"
						+ "WHERE SUBSTR(PROJECT_ID,3)='" + strOldProNo + "' ORDER BY `YEAR`";

				rsTEW2 = insStmtl.executeQuery(sqlselOld);

				if (rsTEW2.next()) {
					// 检索有结果，根据结果创建新的受注（比New能入力更多的项目）
					insertSQL = "replace into T_AOM_NDTL(PROJECT_ID,TEAM,BUSINESS_FIELD,CONTRACTUAL_RELATION,PM_NO,FIELD,WORK_PLACE,`YEAR`,ENT_DT,ENT_ID) values(";

					strVPart1 = "";

					strVPart1 = strVPart1 + "'" + rsTEW2.getString("TEAM") + "',"; // TEAM
					strVPart1 = strVPart1 + "'" + rsTEW2.getString("BUSINESS_FIELD") + "',"; // BUSINESS_FIELD
					strVPart1 = strVPart1 + "'" + rsTEW2.getString("CONTRACTUAL_RELATION") + "',"; // CONTRACTUAL_RELATION
					// *PM_NO仍然用SERP检索出来的（rsTEW），不采用前一年的
					strVPart1 = strVPart1 + "'" + rsTEW.getString("PM_NO") + "',"; // PM_NO
					strVPart1 = strVPart1 + "'" + rsTEW2.getString("FIELD") + "',"; // FIELD

					// strProID1（PROJECT_ID相关）,strVPart2（YEAR,ENT_DT,ENT_ID相关）保持不变
					sql1 = insertSQL + strProID1 + strVPart1 + "'01'," + strVPart2; // JAPAN
					sql2 = insertSQL + strProID2 + strVPart1 + "'02'," + strVPart2; // DALIAN
					sql3 = insertSQL + strProID3 + strVPart1 + "'03'," + strVPart2; // WEIFANG

				}

				// 插入数据（3条：日本，大连，潍坊）
				System.out.println("Inserting into T_AOM_NDTL(3Items)-project_id:XX" + rsTEW.getString("PROJECT_NO")
						+ ", PM_No:" + rsTEW.getString("PM_NO"));
				System.out.println("insert SQL:" + sql1);
				System.out.println("insert SQL:" + sql2);
				System.out.println("insert SQL:" + sql3);

				try {
					insStmtl.executeUpdate(sql1);
					insStmtl.executeUpdate(sql2);
					insStmtl.executeUpdate(sql3);
				} catch (SQLException e) {
					System.out.println("Insert failed!");
					errorcount = errorcount + 1;
					e.printStackTrace();
					throw new Exception(e);
				}

			}

			if (errorcount == 0)
				System.out.println("All New AOMs inserted successfully");
			else
				System.out.println("some or all of new AOMs insert failed");

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
			if (rsTEW2 != null) {
				try {
					rsTEW2.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
				rsTEW2 = null;

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
