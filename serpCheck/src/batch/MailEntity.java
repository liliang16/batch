package batch;

import java.io.Serializable;

/**
* @author liliang
* @since 2020-12-29
*/

public class MailEntity implements Serializable {

    //用来表明类的不同版本间的兼容性
    private static final long serialVersionUID = 1L;
    
    // プロジェクトID
    private String projectId;
    // プロジェクト名称
    private String projectName;
    // PM名字
    private String pmName;
    // 当月小程序人员预订
    private String webNOs;
    // 当月受注管理工数
    private String tAomNOs;
    // 次月SERP人员预订
    private String serpNextNOs;
    // 次月小程序人员预订(0,1,2)
    private String webNextNOs;
    // 次月受注管理工数
    private String tAomNextNOs;
    // 次月小程序人员预订(1,2)
    private String webNextNOs2;
    // PM邮箱
    private String pmMail;
    
	public String getProjectId() {
		return projectId;
	}
	public void setProjectId(String projectId) {
		this.projectId = projectId;
	}
	public String getProjectName() {
		return projectName;
	}
	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}
	public String getPmName() {
		return pmName;
	}
	public void setPmName(String pmName) {
		this.pmName = pmName;
	}
	public String getWebNOs() {
		return webNOs;
	}
	public void setWebNOs(String webNOs) {
		this.webNOs = webNOs;
	}
	public String gettAomNOs() {
		return tAomNOs;
	}
	public void settAomNOs(String tAomNOs) {
		this.tAomNOs = tAomNOs;
	}
	public String getSerpNextNOs() {
		return serpNextNOs;
	}
	public void setSerpNextNOs(String serpNextNOs) {
		this.serpNextNOs = serpNextNOs;
	}
	public String getWebNextNOs() {
		return webNextNOs;
	}
	public void setWebNextNOs(String webNextNOs) {
		this.webNextNOs = webNextNOs;
	}
	public String gettAomNextNOs() {
		return tAomNextNOs;
	}
	public void settAomNextNOs(String tAomNextNOs) {
		this.tAomNextNOs = tAomNextNOs;
	}
	public String getPmMail() {
		return pmMail;
	}
	public void setPmMail(String pmMail) {
		this.pmMail = pmMail;
	}
	public String getWebNextNOs2() {
		return webNextNOs2;
	}
	public void setWebNextNOs2(String webNextNOs2) {
		this.webNextNOs2 = webNextNOs2;
	}
	public static long getSerialversionuid() {
		return serialVersionUID;
	}
}