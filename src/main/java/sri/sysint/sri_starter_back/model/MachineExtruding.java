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
@Table(name = "SRI_IMPP_M_MACHINE_EXTRUDING")
public class MachineExtruding {
	@Id
    @Column(name = "ID_MACHINE_EXT")
    private BigDecimal ID_MACHINE_EXT;
	
	@Column(name = "BUILDING_ID")
    private BigDecimal BUILDING_ID;
	
	@Column(name = "TYPE")
	private String TYPE;
	
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
	
	public MachineExtruding() {
	}
	
	public MachineExtruding(MachineExtruding machineExtruding) {
		this.ID_MACHINE_EXT = machineExtruding.getID_MACHINE_EXT();
		this.BUILDING_ID = machineExtruding.getBUILDING_ID();
		this.TYPE = machineExtruding.getTYPE();
		this.STATUS = machineExtruding.getSTATUS();
		this.CREATION_DATE = machineExtruding.getCREATION_DATE();
		this.CREATED_BY = machineExtruding.getCREATED_BY();
		this.LAST_UPDATE_DATE = machineExtruding.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = machineExtruding.getLAST_UPDATED_BY();
	}
	

	public MachineExtruding(BigDecimal iD_MACHINE_EXT, BigDecimal bUILDING_ID, String tYPE, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		ID_MACHINE_EXT = iD_MACHINE_EXT;
		BUILDING_ID = bUILDING_ID;
		TYPE = tYPE;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getID_MACHINE_EXT() {
		return ID_MACHINE_EXT;
	}

	public void setID_MACHINE_EXT(BigDecimal iD_MACHINE_EXT) {
		ID_MACHINE_EXT = iD_MACHINE_EXT;
	}

	public BigDecimal getBUILDING_ID() {
		return BUILDING_ID;
	}

	public void setBUILDING_ID(BigDecimal bUILDING_ID) {
		BUILDING_ID = bUILDING_ID;
	}

	public String getTYPE() {
		return TYPE;
	}

	public void setTYPE(String tYPE) {
		TYPE = tYPE;
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
