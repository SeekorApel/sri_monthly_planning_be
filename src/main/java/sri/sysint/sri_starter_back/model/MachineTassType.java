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
@Table(name = "SRI_IMPP_M_MACHINETASSTYPE")
public class MachineTassType {
	@Id
    @Column(name = "MACHINETASSTYPE_ID")
    private String MACHINETASSTYPE_ID;
	
	@Column(name = "SETTING_ID")
    private BigDecimal SETTING_ID;
	
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
	
	public MachineTassType() {
	}
	
	public MachineTassType(MachineTassType machineTassType) {
		this.MACHINETASSTYPE_ID = machineTassType.getMACHINETASSTYPE_ID();
        this.SETTING_ID = machineTassType.getSETTING_ID();
        this.DESCRIPTION = machineTassType.getDESCRIPTION();
		this.STATUS = machineTassType.getSTATUS();
		this.CREATION_DATE = machineTassType.getCREATION_DATE();
		this.CREATED_BY = machineTassType.getCREATED_BY();
		this.LAST_UPDATE_DATE = machineTassType.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = machineTassType.getLAST_UPDATED_BY();
	}
	
	

	public MachineTassType(String mACHINETASSTYPE_ID, BigDecimal sETTING_ID, String dESCRIPTION, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		MACHINETASSTYPE_ID = mACHINETASSTYPE_ID;
		SETTING_ID = sETTING_ID;
		DESCRIPTION = dESCRIPTION;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public String getMACHINETASSTYPE_ID() {
		return MACHINETASSTYPE_ID;
	}

	public void setMACHINETASSTYPE_ID(String mACHINETASSTYPE_ID) {
		MACHINETASSTYPE_ID = mACHINETASSTYPE_ID;
	}

	public BigDecimal getSETTING_ID() {
		return SETTING_ID;
	}

	public void setSETTING_ID(BigDecimal sETTING_ID) {
		SETTING_ID = sETTING_ID;
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
