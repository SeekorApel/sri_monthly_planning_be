package sri.sysint.sri_starter_back.model;

import java.io.Serializable;


import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

import sri.sysint.sri_starter_back.util.Auditable;

@Entity
@Table(name = "SRI_APP_USERS")
public class Users extends Auditable<String> implements Serializable {
	
	private static final long serialVersionUID = -8614285921298791389L;
	
	@Id
	@Column(name = "ID")
	private Long id;
	@Column(name="USER_ID")
	private Integer userId;
	@Column(name = "USERNAME")
	private String userName;
	@Column(name = "EMAIL")
	private String email;
	@Column(name = "PASS")
	private String password;
	@Column(name = "PASSCONF")
	private String passWordConf;
	@Column(name = "STATUS")
	private String userStatus;
	@Column(name = "SESSION_KEY")
	private String sessionKey;
	@Column(name = "RESET_TOKEN")
	private String resetToken;
	
	@Column(name = "FULLNAME")
	private String fullname;
		 
	public Users() {
	}
	
	public Users(Users user) {
		this.id = user.getId();
		this.userId = user.getUserId();
		this.userName = user.getUserName();
		this.email = user.getEmail();
		this.password = user.getPassword();
		this.passWordConf = user.getPassWordConf();
		this.userStatus = user.getUserStatus();
		this.sessionKey = user.getSessionKey();
		this.resetToken = user.getResetToken();
		this.fullname = user.getFullname();
	}
	
	public Long getId() {
		return id;
	}
	public void setId(Long id) {
		this.id = id;
	}
	public Integer getUserId() {
		return userId;
	}
	public void setUserId(Integer userId) {
		this.userId = userId;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	
	public String getPassword() {
		return password;
	}
	public void setPassword(String password) {
		this.password = password;
	}
	public String getPassWordConf() {
		return passWordConf;
	}
	public void setPassWordConf(String passWordConf) {
		this.passWordConf = passWordConf;
	}
	public String getUserStatus() {
		return userStatus;
	}
	public void setUserStatus(String userStatus) {
		this.userStatus = userStatus;
	}
	public String getSessionKey() {
		return sessionKey;
	}
	public void setSessionKey(String sessionKey) {
		this.sessionKey = sessionKey;
	}
	public String getResetToken() {
		return resetToken;
	}
	public void setResetToken(String resetToken) {
		this.resetToken = resetToken;
	}

	public String getFullname() {
		return fullname;
	}

	public void setFullname(String fullname) {
		this.fullname = fullname;
	}	
	
	

}
