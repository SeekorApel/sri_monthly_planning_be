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
@Table(name = "SRI_IMPP_M_STOP_MACHINE")
public class StopMachine {
	@Id
    @Column(name = "STOP_MACHINE_ID")
    private BigDecimal STOP_MACHINE_ID;
	@Column(name = "WORK_CENTER_TEXT")
	private String WORK_CENTER_TEXT;
	@Column(name = "DATE_PM")
	private Date DATE_PM;
	@Column(name = "START_TIME")
	private String START_TIME;
	@Column(name = "END_TIME")
	private String END_TIME;
	@Column(name = "TOTAL_TIME")
	private BigDecimal TOTAL_TIME;
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
	
	public StopMachine() {
	}
	
	public StopMachine(StopMachine stopMachine) {
		this.STOP_MACHINE_ID = stopMachine.getSTOP_MACHINE_ID();
		this.WORK_CENTER_TEXT = stopMachine.getWORK_CENTER_TEXT();
		this.DATE_PM = stopMachine.getDATE_PM();
		this.START_TIME = stopMachine.getSTART_TIME();
		this.END_TIME = stopMachine.getEND_TIME();
		this.STATUS = stopMachine.getSTATUS();
		this.CREATION_DATE = stopMachine.getCREATION_DATE();
		this.CREATED_BY = stopMachine.getCREATED_BY();
		this.LAST_UPDATE_DATE = stopMachine.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = stopMachine.getLAST_UPDATED_BY();
	}
	
	

	public StopMachine(BigDecimal sTOP_MACHINE_ID, String wORK_CENTER_TEXT, Date dATE_PM, String sTART_TIME,
			String eND_TIME, BigDecimal tOTAL_TIME, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		STOP_MACHINE_ID = sTOP_MACHINE_ID;
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
		DATE_PM = dATE_PM;
		START_TIME = sTART_TIME;
		END_TIME = eND_TIME;
		TOTAL_TIME = tOTAL_TIME;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getSTOP_MACHINE_ID() {
		return STOP_MACHINE_ID;
	}

	public void setSTOP_MACHINE_ID(BigDecimal sTOP_MACHINE_ID) {
		STOP_MACHINE_ID = sTOP_MACHINE_ID;
	}

	public String getWORK_CENTER_TEXT() {
		return WORK_CENTER_TEXT;
	}

	public void setWORK_CENTER_TEXT(String wORK_CENTER_TEXT) {
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
	}

	public Date getDATE_PM() {
		return DATE_PM;
	}

	public void setDATE_PM(Date dATE_PM) {
		DATE_PM = dATE_PM;
	}

	public String getSTART_TIME() {
		return START_TIME;
	}

	public void setSTART_TIME(String sTART_TIME) {
		START_TIME = sTART_TIME;
	}

	public String getEND_TIME() {
		return END_TIME;
	}

	public void setEND_TIME(String eND_TIME) {
		END_TIME = eND_TIME;
	}

	public BigDecimal getTOTAL_TIME() {
		return TOTAL_TIME;
	}

	public void setTOTAL_TIME(BigDecimal tOTAL_TIME) {
		TOTAL_TIME = tOTAL_TIME;
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
