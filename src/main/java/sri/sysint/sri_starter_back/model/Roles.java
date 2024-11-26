package sri.sysint.sri_starter_back.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_APP_ROLES")
public class Roles {
	
	@Id
	@Column(name = "ID")
	private Long id;
	
	@Column(name = "DESCRIPTION")
	private String description;
	
	@Column(name = "ROLE_NAME")
	private String role_name;
	
	public Roles() {}

	public Roles(Long id, String description, String role_name) {
		super();
		this.id = id;
		this.description = description;
		this.role_name = role_name;
	}
	
	public Roles(Roles roles) {
		this.id = roles.getId();
		this.description = roles.getDescription();
		this.role_name = roles.getRole_name();
	}

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getRole_name() {
		return role_name;
	}

	public void setRole_name(String role_name) {
		this.role_name = role_name;
	}
	
}
