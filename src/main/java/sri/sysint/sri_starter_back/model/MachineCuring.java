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
@Table(name = "SRI_IMPP_M_MACHINE_CURING")
public class MachineCuring {
	@Id
    @Column(name = "WORK_CENTER_TEXT")
    private String WORK_CENTER_TEXT;
	
	@Column(name = "BUILDING_ID")
    private BigDecimal BUILDING_ID;
	
	@Column(name = "CAVITY")
    private BigDecimal CAVITY;
	
    @Column(name = "MACHINE_TYPE")
    private String MACHINE_TYPE;
    
	@Column(name = "STATUS_USAGE")
    private BigDecimal STATUS_USAGE;
    
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
	
	public MachineCuring() {
	}
	
	public MachineCuring(MachineCuring machineCuring) {
		this.WORK_CENTER_TEXT = machineCuring.getWORK_CENTER_TEXT();
		this.BUILDING_ID= machineCuring.getBUILDING_ID();
		this.CAVITY = machineCuring.getCAVITY();
		this.MACHINE_TYPE = machineCuring.getMACHINE_TYPE();
		this.STATUS_USAGE = machineCuring.getSTATUS_USAGE();
		this.STATUS = machineCuring.getSTATUS();
		this.CREATION_DATE = machineCuring.getCREATION_DATE();
		this.CREATED_BY = machineCuring.getCREATED_BY();
		this.LAST_UPDATE_DATE = machineCuring.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = machineCuring.getLAST_UPDATED_BY();
	}


	public MachineCuring(String wORK_CENTER_TEXT, BigDecimal bUILDING_ID,BigDecimal cAVITY,
			String mACHINE_TYPE, BigDecimal sTATUS_USAGE, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
		BUILDING_ID = bUILDING_ID;
		CAVITY = cAVITY;
		MACHINE_TYPE = mACHINE_TYPE;
		STATUS_USAGE = sTATUS_USAGE;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public String getWORK_CENTER_TEXT() {
		return WORK_CENTER_TEXT;
	}

	public void setWORK_CENTER_TEXT(String wORK_CENTER_TEXT) {
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
	}

	public BigDecimal getBUILDING_ID() {
		return BUILDING_ID;
	}

	public void setBUILDING_ID(BigDecimal bUILDING_ID) {
		BUILDING_ID = bUILDING_ID;
	}

	public BigDecimal getCAVITY() {
		return CAVITY;
	}

	public void setCAVITY(BigDecimal cAVITY) {
		CAVITY = cAVITY;
	}

	public String getMACHINE_TYPE() {
		return MACHINE_TYPE;
	}

	public void setMACHINE_TYPE(String mACHINE_TYPE) {
		MACHINE_TYPE = mACHINE_TYPE;
	}

	public BigDecimal getSTATUS_USAGE() {
		return STATUS_USAGE;
	}

	public void setSTATUS_USAGE(BigDecimal sTATUS_USAGE) {
		STATUS_USAGE = sTATUS_USAGE;
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
