package sri.sysint.sri_starter_back.model;

import java.io.Serializable;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

import sri.sysint.sri_starter_back.util.Auditable;

@Entity
@Table(name = "SRI_DUMMIES")
public class Dummy extends Auditable<String> implements Serializable {

	private static final long serialVersionUID = 1601912750015564872L;

	@Id
    @GeneratedValue(strategy =  GenerationType.IDENTITY)
	@Column(name = "ID")
	private Integer id;
	@Column(name = "CODE")
	private String code;
	@Column(name = "DESCRIPTION")
	private String desc;
	
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public String getCode() {
		return code;
	}
	public void setCode(String code) {
		this.code = code;
	}
	public String getDesc() {
		return desc;
	}
	public void setDesc(String desc) {
		this.desc = desc;
	}
}
