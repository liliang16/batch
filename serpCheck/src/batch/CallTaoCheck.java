package batch;

import java.io.IOException;
import java.net.URI;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;

public class CallTaoCheck {

	/** server address */
	static String dbUrl;
	/** MYSQL JDBC driver */
	static String driver;
	/** DB username */
	static String dbUser;
	/** DB password */
	static String dbPassword;
	/** service address */
	static String url;
	/** access user */
	static String user;
	/** wechat WECHAT_ID */
	static String openID;

	public static void main(String[] args) {

		driver = args[0]; // MYSQL JDBC driver
		dbUrl = args[1]; // server address
		dbUser = args[2]; // db username
		dbPassword = args[3].toString(); // db password
		url = args[4]; // web access url
		user = args[5]; // web access user

		System.out.println("DB server address:" + dbUrl);
		System.out.println("MYSQL JDBC driver:" + driver);
		System.out.println("DB username:" + dbUser);
		System.out.println("DB password:" + dbPassword);
		System.out.println("service url:" + url);
		System.out.println("username:" + user);
		/*
		 * // 正式版 // String url
		 * ="http://52.196.8.152:8081/ahiru/taomdtl/saveError"; // 体验版 //
		 * Stringurl = "http://54.248.2.30:8081/ahiruDev/taomdtl/saveError"; //
		 * 本机 //String url = "http://127.0.0.1:8080/ahiru/taomdtl/saveError";
		 * 
		 */
		// 创建Httpclient对象
		CloseableHttpClient httpclient = HttpClients.createDefault();
		String resultString = "";
		CloseableHttpResponse response = null;

		// Load the MYSQL JDBC driver program
		try {
			Class.forName(driver);
			System.out.println("The driver is loaded successfully!");
		} catch (Exception e) {
			System.out.print("The driver failed to load!");
			e.printStackTrace();
			System.exit(1);
		}

		Connection connect = null;
		Statement stmtl = null;
		ResultSet rsTEW = null;

		try {
			// Get OpenID from mysqlDB
			// Database Connection
			connect = DriverManager.getConnection(dbUrl, dbUser, dbPassword);
			System.out.println("The database connection is successful!");

			// select from DB and get openId
			String selUSQL = "SELECT EMPLOYEE_NO, WECHAT_ID FROM M_EMP_DTL where EMPLOYEE_NO = '" + user
					+ "' AND DEL_FG = '0' ";
			stmtl = connect.createStatement();
			rsTEW = stmtl.executeQuery(selUSQL);

			if (rsTEW.next()) {
				openID = rsTEW.getString("WECHAT_ID");
				System.out
						.println("user/openid:" + rsTEW.getString("EMPLOYEE_NO") + "/" + rsTEW.getString("WECHAT_ID"));
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.exit(1);
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
			if (connect != null) {
				try {
					connect.close();
				} catch (Exception e) {
					e.printStackTrace();
				}

			}
		}

		try {
			// 创建uri
			URIBuilder builder = new URIBuilder(url);
			URI uri1 = builder.build();
			HttpGet httpGet = new HttpGet(uri1);

			/*
			 * httpGet.addHeader("USERNAME", "P0000083"); // 体验版
			 * 值：M_EMP_DTL的WECHAT_ID httpGet.addHeader("OPENID",
			 * "oP-ia5a617wKzjfCIZexplQCGcf0"); // 正式版 值：M_EMP_DTL的WECHAT_ID //
			 * httpGet.addHeader("OPENID", "oiN-X5GWde8DPnMhz_AGOpzPlRJk");
			 * 
			 */
			httpGet.addHeader("USERNAME", user);
			httpGet.addHeader("OPENID", openID);
			httpGet.addHeader("Content-type", "application/json; charset=utf-8");

			// 执行请求
			response = httpclient.execute(httpGet);

			// 返回信息
			System.out.println("url" + url);
			System.out.println("getStatusCode" + response.getStatusLine().getStatusCode());

			resultString = EntityUtils.toString(response.getEntity(), "UTF-8");
			System.out.println("resultString" + resultString);

			/*
			 * // 判断返回状态是否为200 if (response.getStatusLine().getStatusCode() ==
			 * 200) { resultString =
			 * EntityUtils.toString(response.getEntity(),"UTF-8");
			 * System.out.println("resultString" + resultString); }
			 */
		} catch (Exception e) {
			System.out.println("systemerror" + e);
		} finally {
			try {
				if (response != null) {
					response.close();
				}
				httpclient.close();
			} catch (IOException e) {
				System.out.println(e);
			}
		}
	}
}
