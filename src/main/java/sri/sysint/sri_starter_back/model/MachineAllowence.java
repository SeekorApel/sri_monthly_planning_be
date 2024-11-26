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
@Table(name = "SRI_IMPP_M_MACHINE_ALLOWENCE")
public class MachineAllowence {
	@Id
    @Column(name = "MACHINE_ALLOW_ID")
    private BigDecimal MACHINE_ALLOW_ID;
	
	@Column(name = "ID_MACHINE")
	private String ID_MACHINE;
	
	@Column(name = "PERSON_RESPONSIBLE")
	private String PERSON_RESPONSIBLE;
	
    @Column(name = "SHIFT_1")
    private BigDecimal SHIFT_1;
    
    @Column(name = "SHIFT_2")
    private BigDecimal SHIFT_2;
    
    @Column(name = "SHIFT_3")
    private BigDecimal SHIFT_3;
    
    @Column(name = "SHIFT_1_FRIDAY")
    private BigDecimal SHIFT_1_FRIDAY;
    
    @Column(name = "TOTAL_SHIFT_123")
    private BigDecimal TOTAL_SHIFT_123;
	
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
	
	public MachineAllowence() {
	}
	
	public MachineAllowence(MachineAllowence machineAllowence) {
		this.MACHINE_ALLOW_ID = machineAllowence.getMACHINE_ALLOW_ID();
		this.ID_MACHINE = machineAllowence.getID_MACHINE();
		this.PERSON_RESPONSIBLE = machineAllowence.getPERSON_RESPONSIBLE();
		this.SHIFT_1 = machineAllowence.getSHIFT_1();
		this.SHIFT_2 = machineAllowence.getSHIFT_2();
		this.SHIFT_3 = machineAllowence.getSHIFT_3();
		this.SHIFT_1_FRIDAY = machineAllowence.getSHIFT_1_FRIDAY();
		this.TOTAL_SHIFT_123 = machineAllowence.getTOTAL_SHIFT_123();
		this.STATUS = machineAllowence.getSTATUS();
		this.CREATION_DATE = machineAllowence.getCREATION_DATE();
		this.CREATED_BY = machineAllowence.getCREATED_BY();
		this.LAST_UPDATE_DATE = machineAllowence.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = machineAllowence.getLAST_UPDATED_BY();
	}
	
	public MachineAllowence(BigDecimal mACHINE_ALLOW_ID, String iD_MACHINE, String pERSON_RESPONSIBLE,
			BigDecimal sHIFT_1, BigDecimal sHIFT_2, BigDecimal sHIFT_3, BigDecimal sHIFT_1_FRIDAY,
			BigDecimal tOTAL_SHIFT_123, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE,
			String lAST_UPDATED_BY) {
		super();
		MACHINE_ALLOW_ID = mACHINE_ALLOW_ID;
		ID_MACHINE = iD_MACHINE;
		PERSON_RESPONSIBLE = pERSON_RESPONSIBLE;
		SHIFT_1 = sHIFT_1;
		SHIFT_2 = sHIFT_2;
		SHIFT_3 = sHIFT_3;
		SHIFT_1_FRIDAY = sHIFT_1_FRIDAY;
		TOTAL_SHIFT_123 = tOTAL_SHIFT_123;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getMACHINE_ALLOW_ID() {
		return MACHINE_ALLOW_ID;
	}

	public void setMACHINE_ALLOW_ID(BigDecimal mACHINE_ALLOW_ID) {
		MACHINE_ALLOW_ID = mACHINE_ALLOW_ID;
	}

	public String getID_MACHINE() {
		return ID_MACHINE;
	}

	public void setID_MACHINE(String iD_MACHINE) {
		ID_MACHINE = iD_MACHINE;
	}

	public String getPERSON_RESPONSIBLE() {
		return PERSON_RESPONSIBLE;
	}

	public void setPERSON_RESPONSIBLE(String pERSON_RESPONSIBLE) {
		PERSON_RESPONSIBLE = pERSON_RESPONSIBLE;
	}

	public BigDecimal getSHIFT_1() {
		return SHIFT_1;
	}

	public void setSHIFT_1(BigDecimal sHIFT_1) {
		SHIFT_1 = sHIFT_1;
	}

	public BigDecimal getSHIFT_2() {
		return SHIFT_2;
	}

	public void setSHIFT_2(BigDecimal sHIFT_2) {
		SHIFT_2 = sHIFT_2;
	}

	public BigDecimal getSHIFT_3() {
		return SHIFT_3;
	}

	public void setSHIFT_3(BigDecimal sHIFT_3) {
		SHIFT_3 = sHIFT_3;
	}

	public BigDecimal getSHIFT_1_FRIDAY() {
		return SHIFT_1_FRIDAY;
	}

	public void setSHIFT_1_FRIDAY(BigDecimal sHIFT_1_FRIDAY) {
		SHIFT_1_FRIDAY = sHIFT_1_FRIDAY;
	}

	public BigDecimal getTOTAL_SHIFT_123() {
		return TOTAL_SHIFT_123;
	}

	public void setTOTAL_SHIFT_123(BigDecimal tOTAL_SHIFT_123) {
		TOTAL_SHIFT_123 = tOTAL_SHIFT_123;
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
