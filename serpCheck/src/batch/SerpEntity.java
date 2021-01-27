package batch;

import java.io.Serializable;

/**
* @author liliang
* @since 2020-07-02
*/

public class SerpEntity implements Serializable {

    //用来表明类的不同版本间的兼容性
    private static final long serialVersionUID = 1L;
    

    // 	プロジェクトID
    private String projectId;
    // 	プロジェクト名
    private String projectName;
    // 	小程序メンバー数
    private String projectMembers;
    // 	serpメンバー数
    private String serpProjectMembers;
    // 	年
    private String projectYear;
    // 	月
    private String projectMonth;
    //  PM名字
    private String pmName;
    //  PMmail
    private String pmmail;
    //  error message
    private String message;
    
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
	public String getProjectMembers() {
		return projectMembers;
	}
	public void setProjectMembers(String projectMembers) {
		this.projectMembers = projectMembers;
	}
	public String getSerpProjectMembers() {
		return serpProjectMembers;
	}
	public void setSerpProjectMembers(String serpProjectMembers) {
		this.serpProjectMembers = serpProjectMembers;
	}
	public String getProjectYear() {
		return projectYear;
	}
	public void setProjectYear(String projectYear) {
		this.projectYear = projectYear;
	}
	public String getProjectMonth() {
		return projectMonth;
	}
	public void setProjectMonth(String projectMonth) {
		this.projectMonth = projectMonth;
	}
	public String getPmName() {
		return pmName;
	}
	public void setPmName(String pmName) {
		this.pmName = pmName;
	}
	public String getPmmail() {
		return pmmail;
	}
	public void setPmmail(String pmmail) {
		this.pmmail = pmmail;
	}
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	public static long getSerialversionuid() {
		return serialVersionUID;
	}
	
}