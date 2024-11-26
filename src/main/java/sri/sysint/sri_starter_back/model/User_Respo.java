package sri.sysint.sri_starter_back.model;

import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_APP_USER_RESPO")
public class User_Respo {
	@Id
	@Column(name = "ID")
	private Long id;
	
	@Column(name = "USER_ID")
	private Long user_id;
	
	@Column(name = "ROLE_ID")
	private Long role_id;
	
	@Column(name = "ACTIVE_DATE")
	private Date active_date;
	
	public User_Respo() {}
	

	public User_Respo(User_Respo user_Respo) {
		this.id = user_Respo.getId();
		this.user_id = user_Respo.getUser_id();
		this.role_id = user_Respo.getRole_id();
		this.active_date = user_Respo.getActive_date();
	}
	
	public User_Respo(Long id, Long user_id, Long role_id, Date active_date) {
		super();
		this.id = id;
		this.user_id = user_id;
		this.role_id = role_id;
		this.active_date = active_date;
	}

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public Long getUser_id() {
		return user_id;
	}

	public void setUser_id(Long user_id) {
		this.user_id = user_id;
	}

	public Long getRole_id() {
		return role_id;
	}

	public void setRole_id(Long role_id) {
		this.role_id = role_id;
	}


	public Date getActive_date() {
		return active_date;
	}


	public void setActive_date(Date active_date) {
		this.active_date = active_date;
	}

	
}
