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
@Table(name = "SRI_IMPP_M_MACHINECURINGTYPE")
public class MachineCuringType {
	@Id
    @Column(name = "MACHINECURINGTYPE_ID")
    private String MACHINECURINGTYPE_ID;
	
	@Column(name = "SETTING_ID")
    private BigDecimal SETTING_ID;
	
	@Column(name = "DESCRIPTION")
	private String DESCRIPTION;
	
	@Column(name = "CAVITY")
    private BigDecimal CAVITY;
	
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
	
	public MachineCuringType() {
	}
	
	public MachineCuringType(MachineCuringType machineCuringType) {
		this.MACHINECURINGTYPE_ID = machineCuringType.getMACHINECURINGTYPE_ID();
        this.SETTING_ID = machineCuringType.getSETTING_ID();
        this.DESCRIPTION = machineCuringType.getDESCRIPTION();
        this.CAVITY = machineCuringType.getCAVITY();
		this.STATUS = machineCuringType.getSTATUS();
		this.CREATION_DATE = machineCuringType.getCREATION_DATE();
		this.CREATED_BY = machineCuringType.getCREATED_BY();
		this.LAST_UPDATE_DATE = machineCuringType.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = machineCuringType.getLAST_UPDATED_BY();
	}
	
	

	public MachineCuringType(String mACHINECURINGTYPE_ID, BigDecimal sETTING_ID, String dESCRIPTION, BigDecimal cAVITY,
			BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		MACHINECURINGTYPE_ID = mACHINECURINGTYPE_ID;
		SETTING_ID = sETTING_ID;
		DESCRIPTION = dESCRIPTION;
		CAVITY = cAVITY;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public String getMACHINECURINGTYPE_ID() {
		return MACHINECURINGTYPE_ID;
	}

	public void setMACHINECURINGTYPE_ID(String mACHINECURINGTYPE_ID) {
		MACHINECURINGTYPE_ID = mACHINECURINGTYPE_ID;
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

	public BigDecimal getCAVITY() {
		return CAVITY;
	}

	public void setCAVITY(BigDecimal cAVITY) {
		CAVITY = cAVITY;
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
