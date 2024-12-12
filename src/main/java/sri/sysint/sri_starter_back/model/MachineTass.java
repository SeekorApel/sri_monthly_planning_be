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
@Table(name = "SRI_IMPP_M_MACHINE_TASS")
public class MachineTass {
	@Id
    @Column(name = "ID_MACHINE_TASS")
    private String ID_MACHINE_TASS;
	
	@Column(name = "BUILDING_ID")
    private BigDecimal BUILDING_ID;
	
	@Column(name = "FLOOR")
    private BigDecimal FLOOR;
	
	@Column(name = "MACHINE_NUMBER")
    private BigDecimal MACHINE_NUMBER;
	
    @Column(name = "MACHINE_TASS_TYPE_ID")
    private String MACHINE_TASS_TYPE_ID;
	
    @Column(name = "WORK_CENTER_TEXT")
    private String WORK_CENTER_TEXT;
    
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
	
	public MachineTass() {
	}
	
	public MachineTass(MachineTass machineTass) {
		this.ID_MACHINE_TASS = machineTass.getID_MACHINE_TASS();
		this.BUILDING_ID = machineTass.getBUILDING_ID();
		this.FLOOR = machineTass.getFLOOR();
		this.MACHINE_NUMBER = machineTass.getMACHINE_NUMBER();
		this.MACHINE_TASS_TYPE_ID = machineTass.getMACHINE_TASS_TYPE_ID();
		this.WORK_CENTER_TEXT = machineTass.getWORK_CENTER_TEXT();
		this.STATUS = machineTass.getSTATUS();
		this.CREATION_DATE = machineTass.getCREATION_DATE();
		this.CREATED_BY = machineTass.getCREATED_BY();
		this.LAST_UPDATE_DATE = machineTass.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = machineTass.getLAST_UPDATED_BY();
	}

	public MachineTass(String iD_MACHINE_TASS, BigDecimal bUILDING_ID, BigDecimal fLOOR, BigDecimal mACHINE_NUMBER,
			String mACHINE_TASS_TYPE_ID, String wORK_CENTER_TEXT, BigDecimal sTATUS, Date cREATION_DATE,
			String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		ID_MACHINE_TASS = iD_MACHINE_TASS;
		BUILDING_ID = bUILDING_ID;
		FLOOR = fLOOR;
		MACHINE_NUMBER = mACHINE_NUMBER;
		MACHINE_TASS_TYPE_ID = mACHINE_TASS_TYPE_ID;
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public String getID_MACHINE_TASS() {
		return ID_MACHINE_TASS;
	}

	public void setID_MACHINE_TASS(String iD_MACHINE_TASS) {
		ID_MACHINE_TASS = iD_MACHINE_TASS;
	}

	public BigDecimal getBUILDING_ID() {
		return BUILDING_ID;
	}

	public void setBUILDING_ID(BigDecimal bUILDING_ID) {
		BUILDING_ID = bUILDING_ID;
	}

	public BigDecimal getFLOOR() {
		return FLOOR;
	}

	public void setFLOOR(BigDecimal fLOOR) {
		FLOOR = fLOOR;
	}

	public BigDecimal getMACHINE_NUMBER() {
		return MACHINE_NUMBER;
	}

	public void setMACHINE_NUMBER(BigDecimal mACHINE_NUMBER) {
		MACHINE_NUMBER = mACHINE_NUMBER;
	}

	public String getMACHINE_TASS_TYPE_ID() {
		return MACHINE_TASS_TYPE_ID;
	}

	public void setMACHINE_TASS_TYPE_ID(String mACHINE_TASS_TYPE_ID) {
		MACHINE_TASS_TYPE_ID = mACHINE_TASS_TYPE_ID;
	}

	public String getWORK_CENTER_TEXT() {
		return WORK_CENTER_TEXT;
	}

	public void setWORK_CENTER_TEXT(String wORK_CENTER_TEXT) {
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
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
