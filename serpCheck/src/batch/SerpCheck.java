
package batch;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URI;
import java.net.URISyntaxException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import javax.mail.MessagingException;
import javax.mail.internet.AddressException;
import javax.mail.internet.MimeUtility;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

/**
 * <p>
 * </p>
 * 
 * @author ll
 * @since 2020-07-12
 */

public class SerpCheck {

	/** MYSQL JDBC driver */
	public static String driver;
	/** server address */
	private static String url;
	/** username */
	private static String user;
	/** password */
	private static String pwd;
	/** emailusername */
	private static String sender;
	/** emailpassword */
	private static String emailPSW;
	private static String host;

	public static void main(String[] args) throws Exception {

		// JDBC 连接数据库信息
		driver = args[0]; // MYSQL JDBC driver
		url = args[1]; // DB server address
		user = args[2]; // DB username
		pwd = args[3]; // DB password
		sender = args[4]; // email username
		emailPSW = args[5]; // email password
		host = args[6]; // email url

		System.out.println("MYSQL JDBC driver: " + driver);
		System.out.println("server address: " + url);
		System.out.println("username: " + user);
		System.out.println("pwd: " + pwd);
		System.out.println("sender: " + sender);
		System.out.println("emailPSW: " + emailPSW);
		System.out.println("host: " + host);

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
			Connection connect = DriverManager.getConnection(url, user, pwd);
			System.out.println("The database connection is successful!");

			Calendar cal = Calendar.getInstance();
			// 获取当月 "年月"
			cal.add(Calendar.MONTH,0);
			int onYear = cal.get(Calendar.YEAR);
			int onMonth = cal.get(Calendar.MONTH) + 1;
			System.out.println("当月sqlYear:" + onYear);
			System.out.println("当月sqlMonth:" + onMonth);
			
			// 获取次月 "年月"
			cal.add(Calendar.MONTH, 1);
			int sqlYear = cal.get(Calendar.YEAR);
			int sqlMonth = cal.get(Calendar.MONTH) + 1;
			System.out.println("次月sqlYear:" + sqlYear);
			System.out.println("次月sqlMonth:" + sqlMonth);

			List<MailEntity> mailEntitys = new ArrayList<MailEntity>();
			List<MailEntity> newMailEntitys = new ArrayList<MailEntity>();
			List<String> aomErrors = new ArrayList<>();
			
			if (onMonth!=12){
				System.out.println("当月为非12月份");
				// 当月和次月 受注管理工数取得
				String onSqlAom = "select PROJECT_ID, case "+onMonth+" when 1 then sum(JAN_NO) when 2 then sum(FEB_NO) when 3 then sum(MAR_NO) "
						+"when 4 then sum(APR_NO) when 5 then sum(MAY_NO) when 6 then sum(JUN_NO) when 7 then sum(JUL_NO) "
						+"when 8 then sum(AUG_NO) when 9 then sum(SEP_NO) when 10 then sum(OCT_NO) when 11 then sum(NOV_NO) "
						+"when 12 then sum(DEC_NO) END as aomNO "
						+ ", case "+sqlMonth+" when 1 then sum(JAN_NO) when 2 then sum(FEB_NO) when 3 then sum(MAR_NO) "
						+"when 4 then sum(APR_NO) when 5 then sum(MAY_NO) when 6 then sum(JUN_NO) when 7 then sum(JUL_NO) "
						+"when 8 then sum(AUG_NO) when 9 then sum(SEP_NO) when 10 then sum(OCT_NO) when 11 then sum(NOV_NO) "
						+"when 12 then sum(DEC_NO) END as aomNextNO "
						+"from T_AOM_NDTL where YEAR = "+onYear+" and DEL_FG = 0 group by  PROJECT_ID";
				System.out.println("受注管理工数取得(当月和次月)sql:" +onSqlAom);
				
				Statement stmt = connect.createStatement();
				ResultSet onAom = stmt.executeQuery(onSqlAom);
				while (onAom.next()) {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(onAom.getString("PROJECT_ID"));
					if(onAom.getString("aomNO")!=null){
						mailEntity.settAomNOs(onAom.getString("aomNO").replace(".0", ""));
					}
					if(onAom.getString("aomNextNO")!=null){
						mailEntity.settAomNextNOs(onAom.getString("aomNextNO").replace(".0", ""));
					}
					mailEntitys.add(mailEntity);
					aomErrors.add(onAom.getString("PROJECT_ID"));
				}
				
			} else {
				System.out.println("当月为12月份");
				
				// 当月受注管理工数取得
				String onSqlAom = "select PROJECT_ID, sum(DEC_NO) as aomNO from T_AOM_NDTL where YEAR = "+onYear+" and DEL_FG = 0 group by  PROJECT_ID";
				System.out.println("当月受注管理工数取得:" +onSqlAom);
				
				Statement stmt = connect.createStatement();
				ResultSet onAom = stmt.executeQuery(onSqlAom);
				while (onAom.next()) {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(onAom.getString("PROJECT_ID"));
					if(onAom.getString("aomNO")!=null){
						mailEntity.settAomNOs(onAom.getString("aomNO").replace(".0", ""));
					}
					mailEntitys.add(mailEntity);
					aomErrors.add(onAom.getString("PROJECT_ID"));
				}
				
				// 次月受注管理工数取得
				String sqlAomNext = "select PROJECT_ID, sum(JAN_NO) as aomNextNO from T_AOM_NDTL TAN where YEAR = "+sqlYear+" and DEL_FG = 0 group by PROJECT_ID";
				System.out.println("次月受注管理工数取得:" +sqlAomNext);
				Statement stmt1 = connect.createStatement();
				ResultSet onAomNext = stmt1.executeQuery(sqlAomNext);
				while (onAomNext.next()) {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(onAomNext.getString("PROJECT_ID"));
					if(onAomNext.getString("aomNextNO")!=null){
						mailEntity.settAomNextNOs(onAomNext.getString("aomNextNO").replace(".0", ""));
					}
					mailEntitys.add(mailEntity);
					aomErrors.add(onAomNext.getString("PROJECT_ID"));
				}
				
			}
			
			// 当月小程序人员数量取得
			// 不包含backup人员(1,2)
			String sqlTEmpWorkNo = "SELECT PROJECT_ID, COUNT(PROJECT_ID) as Pno FROM T_EMP_WORK WHERE DEL_FG = '0' "
					+"AND USE_STATUS in (1, 2) AND YEAR = "+onYear+" AND USE_MONTH = "+onMonth+" GROUP BY PROJECT_ID";
			System.out.println("当月小程序人员数量取得sql: " +sqlTEmpWorkNo);
			Statement stmt2 = connect.createStatement();
			ResultSet rsTEW = stmt2.executeQuery(sqlTEmpWorkNo);
			while (rsTEW.next()) {
				if (aomErrors.contains(rsTEW.getString("PROJECT_ID"))) {
					for (int i = 0; i < mailEntitys.size(); i++) {
						if (mailEntitys.get(i).getProjectId().equals(rsTEW.getString("PROJECT_ID"))) {
							mailEntitys.get(i).setWebNOs(rsTEW.getString("Pno"));
						}
					}
				} else {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(rsTEW.getString("PROJECT_ID"));
					mailEntity.setWebNOs(rsTEW.getString("Pno"));
					mailEntitys.add(mailEntity);
					aomErrors.add(rsTEW.getString("PROJECT_ID"));
				}
			}
			// 次月小程序人员数量取得
			// 包含backup人员(0,1,2)
			String sqlTEmpWorkNextNo = "SELECT PROJECT_ID, COUNT(PROJECT_ID) as Pno FROM T_EMP_WORK WHERE DEL_FG = '0' "
					+"AND USE_STATUS in (0, 1, 2) AND YEAR = "+sqlYear+" AND USE_MONTH = "+sqlMonth+" GROUP BY PROJECT_ID";
			System.out.println("次月小程序人员数量取得sql: " +sqlTEmpWorkNextNo);
			Statement stmt3 = connect.createStatement();
			ResultSet rsNextTEW = stmt3.executeQuery(sqlTEmpWorkNextNo);
			while (rsNextTEW.next()) {
				if (aomErrors.contains(rsNextTEW.getString("PROJECT_ID"))) {
					for (int i = 0; i < mailEntitys.size(); i++) {
						if (mailEntitys.get(i).getProjectId().equals(rsNextTEW.getString("PROJECT_ID"))) {
							mailEntitys.get(i).setWebNextNOs(rsNextTEW.getString("Pno"));
						}
					}
				} else {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(rsNextTEW.getString("PROJECT_ID"));
					mailEntity.setWebNextNOs(rsNextTEW.getString("Pno"));
					mailEntitys.add(mailEntity);
					aomErrors.add(rsNextTEW.getString("PROJECT_ID"));
				}
			}
			// 次月小程序人员数量取得
			// 不包含backup人员(1,2)
			String sqlTEmpWorkNextNo2 = "SELECT PROJECT_ID, COUNT(PROJECT_ID) as Pno FROM T_EMP_WORK WHERE DEL_FG = '0' "
					+"AND USE_STATUS in (1, 2) AND YEAR = "+sqlYear+" AND USE_MONTH = "+sqlMonth+" GROUP BY PROJECT_ID";
			System.out.println("次月小程序人员数量取得sql: " +sqlTEmpWorkNextNo2);
			Statement stmt4 = connect.createStatement();
			ResultSet rsNextTEW2 = stmt4.executeQuery(sqlTEmpWorkNextNo2);
			while (rsNextTEW2.next()) {
				if (aomErrors.contains(rsNextTEW2.getString("PROJECT_ID"))) {
					for (int i = 0; i < mailEntitys.size(); i++) {
						if (mailEntitys.get(i).getProjectId().equals(rsNextTEW2.getString("PROJECT_ID"))) {
							mailEntitys.get(i).setWebNextNOs2(rsNextTEW2.getString("Pno"));
						}
					}
				} else {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(rsNextTEW2.getString("PROJECT_ID"));
					mailEntity.setWebNextNOs2(rsNextTEW2.getString("Pno"));
					mailEntitys.add(mailEntity);
					aomErrors.add(rsNextTEW2.getString("PROJECT_ID"));
				}
			}
			
			// 次月SERP人员数量取得
			String sqlTSERPDATANo = "SELECT PROJECT_ID, COUNT(PROJECT_ID) as Pno FROM T_SERP_DATA WHERE DEL_FG = '0' "
					+"AND COST_MANHOUR > 0 AND PROJECT_YEAR = "+sqlYear+" AND PROJECT_MONTH = "+ sqlMonth+" GROUP BY PROJECT_ID";
			System.out.println("次月SERP人员数量取得sql: " + sqlTSERPDATANo);

			Statement stmt5 = connect.createStatement();
			ResultSet rsTSERPDATANo = stmt5.executeQuery(sqlTSERPDATANo);
			while (rsTSERPDATANo.next()) {
				if (aomErrors.contains(rsTSERPDATANo.getString("PROJECT_ID"))) {
					for (int i = 0; i < mailEntitys.size(); i++) {
						if (mailEntitys.get(i).getProjectId().equals(rsTSERPDATANo.getString("PROJECT_ID"))) {
							mailEntitys.get(i).setSerpNextNOs(rsTSERPDATANo.getString("Pno"));
						}
					}
				} else {
					MailEntity mailEntity = new MailEntity();
					mailEntity.setProjectId(rsTSERPDATANo.getString("PROJECT_ID"));
					mailEntity.setSerpNextNOs(rsTSERPDATANo.getString("Pno"));
					mailEntitys.add(mailEntity);
					aomErrors.add(rsTSERPDATANo.getString("PROJECT_ID"));
				}
			}
			
			// 去除工数相等的项目
			System.out.println("去除工数相等的项目");
			for (int i = 0; i < mailEntitys.size(); i++) {
				// 清空当月正常项目数据
				// 受注管理工数 VS 小程序项目（受注比较）
				if (mailEntitys.get(i).gettAomNOs()!=null && mailEntitys.get(i).getWebNOs()!=null 
						&& mailEntitys.get(i).gettAomNOs().equals(mailEntitys.get(i).getWebNOs()) ){
					mailEntitys.get(i).settAomNOs("");
					mailEntitys.get(i).setWebNOs("");
				} else if (mailEntitys.get(i).gettAomNOs()!=null && mailEntitys.get(i).getWebNOs()==null){
					mailEntitys.get(i).setWebNOs("0");
				} else if (mailEntitys.get(i).gettAomNOs()==null && mailEntitys.get(i).getWebNOs()!=null){
					mailEntitys.get(i).settAomNOs("0");
				}
				// 清空次月正常项目数据
				// SERP人员预订 VS 小程序项目（SERP比较）
				if (mailEntitys.get(i).getWebNextNOs()!=null && mailEntitys.get(i).getSerpNextNOs()!=null 
						&& mailEntitys.get(i).getSerpNextNOs().equals(mailEntitys.get(i).getWebNextNOs())){
					mailEntitys.get(i).setWebNextNOs("");
					mailEntitys.get(i).setSerpNextNOs("");
				} else if (mailEntitys.get(i).getWebNextNOs()!=null && mailEntitys.get(i).getSerpNextNOs()==null){
					mailEntitys.get(i).setSerpNextNOs("0");
				} else if (mailEntitys.get(i).getWebNextNOs()==null && mailEntitys.get(i).getSerpNextNOs()!=null){
					mailEntitys.get(i).setWebNextNOs("0");
				}
				// 受注管理工数 VS 小程序项目（受注比较）
				if (mailEntitys.get(i).gettAomNextNOs()!=null && mailEntitys.get(i).getWebNextNOs2()!=null
						&& mailEntitys.get(i).gettAomNextNOs().equals(mailEntitys.get(i).getWebNextNOs2())){
					mailEntitys.get(i).settAomNextNOs("");
					mailEntitys.get(i).setWebNextNOs2("");
				} else if (mailEntitys.get(i).gettAomNextNOs()!=null && mailEntitys.get(i).getWebNextNOs2()==null){
					mailEntitys.get(i).setWebNextNOs2("0");
				} else if (mailEntitys.get(i).gettAomNextNOs()==null && mailEntitys.get(i).getWebNextNOs2()!=null){
					mailEntitys.get(i).settAomNextNOs("0");
				}
				if((mailEntitys.get(i).getWebNOs()==null || "".endsWith(mailEntitys.get(i).getWebNOs())) 
						&& (mailEntitys.get(i).getWebNextNOs()==null || "".equals(mailEntitys.get(i).getWebNextNOs()))
						&& (mailEntitys.get(i).gettAomNOs()==null || "".endsWith(mailEntitys.get(i).gettAomNOs())) 
						&& (mailEntitys.get(i).gettAomNextNOs()==null || "".equals(mailEntitys.get(i).gettAomNextNOs()))
						&& (mailEntitys.get(i).getSerpNextNOs()==null || "".equals(mailEntitys.get(i).getSerpNextNOs()))
						&& (mailEntitys.get(i).getWebNextNOs2()==null || "".equals(mailEntitys.get(i).getWebNextNOs2()))){
				}else{
					newMailEntitys.add(mailEntitys.get(i));
				}
			}
			
			// 补全项目信息
			for (int i = 0; i < newMailEntitys.size(); i++) {
				String mpsql = "SELECT MP.PROJECT_NAME, ME.NAME, ME.EMAIL FROM M_PROJECT MP inner join M_EMPLOYEE ME on ME.EMPLOYEE_NO = MP.PM_EMPLOYEE_NO "
							+ "and ME.DEL_FG = '0' WHERE MP.DEL_FG = '0' AND MP.PROJECT_ID = '"+newMailEntitys.get(i).getProjectId()+"'";
				System.out.println("补全项目信息sql: " + mpsql);

				Statement stmt6 = connect.createStatement();
				ResultSet rsMP = stmt6.executeQuery(mpsql);
				while (rsMP.next()) {
					newMailEntitys.get(i).setProjectName(rsMP.getString("PROJECT_NAME"));
					newMailEntitys.get(i).setPmName(rsMP.getString("NAME"));
					newMailEntitys.get(i).setPmMail(rsMP.getString("EMAIL"));
				}
			}

			// 去除重复邮箱地址
			List<String> pmMails = new ArrayList<String>();
			for (MailEntity mailEntity : newMailEntitys) {
				if (!pmMails.contains(mailEntity.getPmMail())) {
					pmMails.add(mailEntity.getPmMail());
				}
			}

			// 给PM发错误邮件提醒
			if (mailEntitys.size()!=0){
				sendMail(pmMails, newMailEntitys);
			}

		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception(e);
		} finally {
			System.out.println("End of all processing");
			System.exit(0);
		}
	}

	public static void sendMail(List<String> to, List<MailEntity> mailEntitys) {

		// 收件人电子邮箱
		List<String> cc = new ArrayList<>();
		String[] receiverCC = getSerpProperty("EmailReceiverCC").split(",");
		for (String receiver :receiverCC) {
			cc.add(receiver);
		}
		
		// TODO
		to = new ArrayList<>();
		to.add("liang.li16@pactera.com");
//		to.add("hongye.su@pactera.com");
		cc = new ArrayList<>();
		cc.add("liang.li16@pactera.com");
//		cc.add("hongye.su@pactera.com");
		
		String subject = "SERP人员工数和受注异常提示";
		String content = "Dear<br><br>&nbsp&nbsp&nbsp&nbspSERP人员工数和受注异常，已生成Excel文件，请查收!<br><br>以上。";
		
		List<String> filePath = new ArrayList<>();
		
		String path = getSerpProperty("filePath") + getNowTime() + getSerpProperty("fileType");
		File file = new File(path);
		try {
			EmailMessage message = new EmailMessage(getExchangeService());
			creatExcel(path, mailEntitys);
			filePath.add(path);
			// 设置邮件主题
			message.setSubject(subject);

			// 设置邮件内容
			MessageBody body = MessageBody.getMessageBodyFromText(content);
			body.setBodyType(BodyType.HTML);
			message.setBody(body);

			// 设置收件人
			for (String address : to) {
				message.getToRecipients().add(address);
			}

			// 设置抄送
			for (String address : cc) {
				message.getCcRecipients().add(address);
			}

			// 设置附件
			for (String file1 :filePath) {
				message.getAttachments().addFileAttachment(file1); 
			}

			// 发送消息
			message.send();
			System.out.println("SendMail Successfully");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (URISyntaxException e) {
			e.printStackTrace();
		} catch (Throwable e) {
			e.printStackTrace();
		} finally {
			file.delete();
			System.out.println(getSerpProperty("filePath") + getNowTime() + getSerpProperty("fileType") +" delete!");
		}
	
	}

	private static String getNowTime() {
		Date date = new Date();
		DateFormat format = new SimpleDateFormat("yyyyMMdd");
		return format.format(date);
	}

	// Excel表格做成
	// Excel第1行（标题行）定义
	private static String[] titleArr =
			{"ID", "名称", "PM", "受注管理工数", "小程序项目（受注比较）", "SERP人员预订", "小程序项目（SERP比较）", "受注管理工数", "小程序项目（受注比较）"};

	private static HSSFCellStyle styleDetailUp;

	private static HSSFCellStyle styleDetailDown;
	// Title的样式样式
	private static HSSFCellStyle styleTitle;
	// Excel的样式样式
	private static HSSFCellStyle styleTitle1;
	private static HSSFCellStyle styleTitle2;
	private static HSSFCellStyle styleTitle3;
	// 右侧单元格样式(上)
	private static HSSFCellStyle styleRightCellUp;
	// 右侧单元格样式（下）
	private static HSSFCellStyle styleRightCellDown;

	private static int creatExcel(String filePath, List<MailEntity> mailEntitys) throws FileNotFoundException {
		// 建立book
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 建立样式
		creatStyle(workbook);
		// 建立sheet
		HSSFSheet sheet = workbook.createSheet();
		// 创建sheet名称
		workbook.setSheetName(0, "SERP小程序受注异常");
		creatRow1(workbook, sheet);
		creatDetail(workbook, sheet, mailEntitys);
		// Excel表格做成
		FileOutputStream fileOut = new FileOutputStream(filePath);
		try {
			workbook.write(fileOut);
		} catch (IOException e) {
			return 0;
		} finally {
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return 1;
	}

	// Excel第1行做成
	private static void creatRow1(HSSFWorkbook workbook, HSSFSheet sheet) {
		// 获取当前年和下个月
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.MONTH, 0);
		Integer sqlYear = cal.get(Calendar.YEAR);
		Integer sqlMonth = cal.get(Calendar.MONTH) + 1;
		
		// 获取当前年和下个月
		Calendar cal2 = Calendar.getInstance();
		cal2.add(Calendar.MONTH, 1);
		Integer sqlYear2 = cal2.get(Calendar.YEAR);
		Integer sqlMonth2 = cal2.get(Calendar.MONTH) + 1;
		// 添加第1行(标题行)
		HSSFRow row0 = sheet.createRow(0);
		HSSFRow row1 = sheet.createRow(1);

        //当前月
		HSSFCell cell0 = row0.createCell(3);
		cell0.setCellValue(sqlYear+"年"+sqlMonth+"月");
		cell0.setCellStyle(styleTitle);
		//次月
		HSSFCell cell1 = row0.createCell(5);
		cell1.setCellValue(sqlYear2+"年"+sqlMonth2+"月");
		cell1.setCellStyle(styleTitle);
		
		// 设置Excel表头
		for (int i = 0; i < titleArr.length; i++) {
			HSSFCell cell = row1.createCell(i);
			cell.setCellValue(titleArr[i]);
			cell.setCellStyle(styleTitle);
			//sheet.autoSizeColumn((short) 3);
		}
		
		// 設置寬度
		sheet.setColumnWidth(0, (int) ((18 + 0.72) * 256));
		sheet.setColumnWidth(1, (int) ((50 + 0.72) * 256));
		sheet.setColumnWidth(2, (int) ((8 + 0.72) * 256));
		sheet.setColumnWidth(3, (int) ((16 + 0.72) * 256));
		sheet.setColumnWidth(4, (int) ((22 + 0.72) * 256));
		sheet.setColumnWidth(5, (int) ((16 + 0.72) * 256));
		sheet.setColumnWidth(6, (int) ((22 + 0.72) * 256));
		sheet.setColumnWidth(7, (int) ((16 + 0.72) * 256));
		sheet.setColumnWidth(8, (int) ((22 + 0.72) * 256));
		
		// 合并单元格 (4个参数，分别为起始行，结束行，起始列，结束列)
        // 行和列都是从0开始计数，且起始结束都会合并
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 4));	//当前年月
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 8));	//次月
	}

	// Excel明细行做成
	private static void creatDetail(HSSFWorkbook workbook, HSSFSheet sheet, List<MailEntity> mailEntitys) {
		int rowNo = 1;
		HSSFRow row = null;
		DecimalFormat df = new DecimalFormat("#.0");
		for (int j = 0; j < mailEntitys.size(); j++) {
			rowNo++;
			row = sheet.createRow(rowNo);
			int colNo = 0;
			// 设置单元格样式
			for (int i = 0; i < titleArr.length; i++) {
				row.createCell(i);
				row.getCell(i).setCellStyle(styleDetailUp);
				row.getCell(i).setCellStyle(styleDetailDown);
				if (i==3 || i==4){
					row.getCell(i).setCellStyle(styleTitle1);
				} else if (i==5 || i==6){
					row.getCell(i).setCellStyle(styleTitle2);
				} else if (i==7 || i==8){
					row.getCell(i).setCellStyle(styleTitle3);
				}
			}
			// 给单元格赋值（员工基本信息）
			row.getCell(colNo++).setCellValue(mailEntitys.get(j).getProjectId());
			row.getCell(colNo++).setCellValue(mailEntitys.get(j).getProjectName());
			row.getCell(colNo++).setCellValue(mailEntitys.get(j).getPmName());
			if (mailEntitys.get(j).gettAomNOs()!=null && mailEntitys.get(j).gettAomNOs()!=""){
				row.getCell(colNo++).setCellValue(Double.valueOf(mailEntitys.get(j).gettAomNOs()));
			} else {
				colNo++;
			}
			if (mailEntitys.get(j).getWebNOs()!=null && mailEntitys.get(j).getWebNOs()!=""){
				row.getCell(colNo++).setCellValue(Double.valueOf(mailEntitys.get(j).getWebNOs()));
			} else {
				colNo++;
			}
			if (mailEntitys.get(j).getSerpNextNOs()!=null && mailEntitys.get(j).getSerpNextNOs()!=""){
				row.getCell(colNo++).setCellValue(Double.valueOf(mailEntitys.get(j).getSerpNextNOs()));
			} else {
				colNo++;
			}
			if (mailEntitys.get(j).getWebNextNOs()!=null && mailEntitys.get(j).getWebNextNOs()!=""){
				row.getCell(colNo++).setCellValue(Double.valueOf(mailEntitys.get(j).getWebNextNOs()));
			} else {
				colNo++;
			}
			if (mailEntitys.get(j).gettAomNextNOs()!=null && mailEntitys.get(j).gettAomNextNOs()!=""){
				row.getCell(colNo++).setCellValue(Double.valueOf(mailEntitys.get(j).gettAomNextNOs()));
			} else {
				colNo++;
			}
			if (mailEntitys.get(j).getWebNextNOs2()!=null && mailEntitys.get(j).getWebNextNOs2()!=""){
				row.getCell(colNo++).setCellValue(Double.valueOf(mailEntitys.get(j).getWebNextNOs2()));
			} else {
				colNo++;
			}
		}
	}

	// 设置样式
	private static void creatStyle(HSSFWorkbook workbook) {
		creatDetailCellFontUp(workbook);
		creatDetailCellFontDown(workbook);
		creatTitleStyle(workbook);
		creatRightCellFontUp(workbook);
		creatRightCellFontDown(workbook);
		creatLeftCellFont(workbook);
		creatBottomCellFont(workbook);
		creatLeftBottomCellFont(workbook);
		creatRightBottomCellFont(workbook);
		// 受注管理工数/小程序项目（受注比较）/SERP人员预订/小程序项目（SERP比较）/受注管理工数/小程序项目（受注比较）  样式
		excelStyle(workbook);
	}

	// 标题行单元格样式
	private static void creatTitleStyle(HSSFWorkbook workbook) {
		// 创建单元格
		styleTitle = workbook.createCellStyle();
		// 设置标题行格式（粗体）
		styleTitle.setAlignment(HorizontalAlignment.CENTER);
		HSSFFont font = workbook.createFont();
		font.setBold(true);
		styleTitle.setFont(font);
		// 设置边框
		styleTitle.setBorderBottom(BorderStyle.THIN);
		styleTitle.setBorderLeft(BorderStyle.THIN);
		styleTitle.setBorderRight(BorderStyle.THIN);
		styleTitle.setBorderTop(BorderStyle.THIN);
		// 背景色浅灰
		styleTitle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		// 填充背景色
		styleTitle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	// Excel明细行单元格样式
	private static void excelStyle(HSSFWorkbook workbook) {
		// 创建单元格
		styleTitle1 = workbook.createCellStyle();
		// 设置边框
		styleTitle1.setBorderBottom(BorderStyle.THIN);
		styleTitle1.setBorderLeft(BorderStyle.THIN);
		styleTitle1.setBorderRight(BorderStyle.THIN);
		styleTitle1.setBorderTop(BorderStyle.THIN);
		// 设置背景色
		styleTitle1.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
		styleTitle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		// 创建单元格
		styleTitle2 = workbook.createCellStyle();
		// 设置边框
		styleTitle2.setBorderBottom(BorderStyle.THIN);
		styleTitle2.setBorderLeft(BorderStyle.THIN);
		styleTitle2.setBorderRight(BorderStyle.THIN);
		styleTitle2.setBorderTop(BorderStyle.THIN);
		// 设置背景色
		styleTitle2.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		styleTitle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// 创建单元格
		styleTitle3 = workbook.createCellStyle();
		// 设置边框
		styleTitle3.setBorderBottom(BorderStyle.THIN);
		styleTitle3.setBorderLeft(BorderStyle.THIN);
		styleTitle3.setBorderRight(BorderStyle.THIN);
		styleTitle3.setBorderTop(BorderStyle.THIN);
		// 设置背景色
		styleTitle3.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
		styleTitle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}
	// 左侧单元格样式
	private static HSSFCellStyle creatLeftCellFont(HSSFWorkbook workbook) {
		HSSFCellStyle styleLeftCell = workbook.createCellStyle();
		styleLeftCell.setBorderBottom(BorderStyle.DASHED);
		styleLeftCell.setBorderLeft(BorderStyle.THIN);
		styleLeftCell.setBorderRight(BorderStyle.DASHED);
		styleLeftCell.setBorderTop(BorderStyle.DASHED);
		return styleLeftCell;
	}

	// 右侧单元格样式
	private static HSSFCellStyle creatRightCellFontUp(HSSFWorkbook workbook) {
		styleRightCellUp = workbook.createCellStyle();
		styleRightCellUp.setBorderLeft(BorderStyle.DASHED);
		styleRightCellUp.setBorderRight(BorderStyle.THIN);
		styleRightCellUp.setBorderTop(BorderStyle.THIN);
		return styleRightCellUp;
	}

	// 右侧单元格样式
	private static HSSFCellStyle creatRightCellFontDown(HSSFWorkbook workbook) {
		styleRightCellDown = workbook.createCellStyle();
		styleRightCellDown.setBorderLeft(BorderStyle.DASHED);
		styleRightCellDown.setBorderRight(BorderStyle.THIN);
		styleRightCellDown.setBorderBottom(BorderStyle.THIN);
		return styleRightCellDown;
	}

	// 下侧单元格样式
	private static HSSFCellStyle creatBottomCellFont(HSSFWorkbook workbook) {
		HSSFCellStyle styleLeftCell = workbook.createCellStyle();
		styleLeftCell.setBorderTop(BorderStyle.DASHED);
		styleLeftCell.setBorderLeft(BorderStyle.DASHED);
		styleLeftCell.setBorderRight(BorderStyle.DASHED);
		styleLeftCell.setBorderBottom(BorderStyle.THIN);
		return styleLeftCell;
	}

	// 左下侧单元格样式
	private static HSSFCellStyle creatLeftBottomCellFont(HSSFWorkbook workbook) {
		HSSFCellStyle styleLeftCell = workbook.createCellStyle();
		styleLeftCell.setBorderTop(BorderStyle.DASHED);
		styleLeftCell.setBorderLeft(BorderStyle.THIN);
		styleLeftCell.setBorderRight(BorderStyle.DASHED);
		styleLeftCell.setBorderBottom(BorderStyle.THIN);
		return styleLeftCell;
	}

	// 右下侧单元格样式
	private static HSSFCellStyle creatRightBottomCellFont(HSSFWorkbook workbook) {
		HSSFCellStyle styleLeftCell = workbook.createCellStyle();
		styleLeftCell.setBorderTop(BorderStyle.DASHED);
		styleLeftCell.setBorderLeft(BorderStyle.DASHED);
		styleLeftCell.setBorderRight(BorderStyle.THIN);
		styleLeftCell.setBorderBottom(BorderStyle.THIN);
		return styleLeftCell;
	}

	// 明细单元格上行样式
	private static void creatDetailCellFontUp(HSSFWorkbook workbook) {
		styleDetailUp = workbook.createCellStyle();
		styleDetailUp.setBorderLeft(BorderStyle.DASHED);
		styleDetailUp.setBorderRight(BorderStyle.DASHED);
		styleDetailUp.setBorderTop(BorderStyle.THIN);
	}

	// 明细单元格下行样式
	private static void creatDetailCellFontDown(HSSFWorkbook workbook) {
		styleDetailDown = workbook.createCellStyle();
		styleDetailDown.setBorderBottom(BorderStyle.THIN);
		styleDetailDown.setBorderLeft(BorderStyle.DASHED);
		styleDetailDown.setBorderRight(BorderStyle.DASHED);

	}
	
	
	/**
	 * 取得excel.properties的值
	 */
	public static String getSerpProperty(String key) {
		final ResourceBundle bundle = ResourceBundle.getBundle("excel");
		return bundle.getString(key);
	}

	/**
	 * 邮箱配置
	 * @return
	 * @throws URISyntaxException
	 */
	private static ExchangeService getExchangeService() throws URISyntaxException {
		ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials(sender, emailPSW);
		exchangeService.setCredentials(credentials);
		exchangeService.setUrl(new URI(host));
		exchangeService.setTimeout(20 * 1000);
		return exchangeService;
	}
}