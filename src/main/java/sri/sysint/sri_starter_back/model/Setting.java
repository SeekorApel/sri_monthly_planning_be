package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

import com.fasterxml.jackson.annotation.JsonProperty;

@Entity
@Table(name = "SRI_IMPP_M_SETTING")
public class Setting {
	@Id
    @Column(name = "SETTING_ID")
    private BigDecimal SETTING_ID;
	
	@Column(name = "SETTING_KEY")
	private String SETTING_KEY;
	
	@Column(name = "SETTING_VALUE")
	private String SETTING_VALUE;
	
	@Column(name = "DESCRIPTION")
	private String DESCRIPTION;
	
	@Column(name = "STATUS")
	private BigDecimal STATUS;
	
	@Column(name = "CREATION_DATE")
	private Date CREATION_DATE;
	
	@Column(name = "CREATED_BY")
	private String CREATED_BY;
	
	@Column(name = "LAST_UPDATE_DATE")
	private Date LAST_UPDATE_DATE;
	
	@Column(name = "LAST_UPDATED_BY")
	private String LAST_UPDATED_BY;
	
	public Setting() {
	}
	
	public Setting(Setting setting) {
        this.SETTING_ID = setting.SETTING_ID;
        this.SETTING_KEY = setting.SETTING_KEY;
        this.SETTING_VALUE = setting.SETTING_VALUE;
        this.DESCRIPTION = setting.DESCRIPTION;
		this.STATUS = setting.getSTATUS();
		this.CREATION_DATE = setting.getCREATION_DATE();
		this.CREATED_BY = setting.getCREATED_BY();
		this.LAST_UPDATE_DATE = setting.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = setting.getLAST_UPDATED_BY();
	}

	public Setting(BigDecimal sETTING_ID, String sETTING_KEY, String sETTING_VALUE, String dESCRIPTION,
			BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		SETTING_ID = sETTING_ID;
		SETTING_KEY = sETTING_KEY;
		SETTING_VALUE = sETTING_VALUE;
		DESCRIPTION = dESCRIPTION;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getSETTING_ID() {
		return SETTING_ID;
	}

	public void setSETTING_ID(BigDecimal sETTING_ID) {
		SETTING_ID = sETTING_ID;
	}

	public String getSETTING_KEY() {
		return SETTING_KEY;
	}

	public void setSETTING_KEY(String sETTING_KEY) {
		SETTING_KEY = sETTING_KEY;
	}

	public String getSETTING_VALUE() {
		return SETTING_VALUE;
	}

	public void setSETTING_VALUE(String sETTING_VALUE) {
		SETTING_VALUE = sETTING_VALUE;
	}

	public String getDESCRIPTION() {
		return DESCRIPTION;
	}

	public void setDESCRIPTION(String dESCRIPTION) {
		DESCRIPTION = dESCRIPTION;
	}

	public BigDecimal getSTATUS() {
		return STATUS;
	}

	public void setSTATUS(BigDecimal sTATUS) {
		STATUS = sTATUS;
	}

	public Date getCREATION_DATE() {
		return CREATION_DATE;
	}

	public void setCREATION_DATE(Date cREATION_DATE) {
		CREATION_DATE = cREATION_DATE;
	}

	public String getCREATED_BY() {
		return CREATED_BY;
	}

	public void setCREATED_BY(String cREATED_BY) {
		CREATED_BY = cREATED_BY;
	}

	public Date getLAST_UPDATE_DATE() {
		return LAST_UPDATE_DATE;
	}

	public void setLAST_UPDATE_DATE(Date lAST_UPDATE_DATE) {
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
	}

	public String getLAST_UPDATED_BY() {
		return LAST_UPDATED_BY;
	}

	public void setLAST_UPDATED_BY(String lAST_UPDATED_BY) {
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}
	
	
}
