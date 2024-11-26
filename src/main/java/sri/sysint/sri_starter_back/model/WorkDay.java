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
@Table(name = "SRI_IMPP_M_WD")
public class WorkDay {
	@Id
    @Column(name = "DATE_WD")
    private Date DATE_WD;
	
	@Column(name = "IWD_SHIFT_1")
	private BigDecimal IWD_SHIFT_1;
	
	@Column(name = "IWD_SHIFT_2")
	private BigDecimal IWD_SHIFT_2;
	
	@Column(name = "IWD_SHIFT_3")
	private BigDecimal IWD_SHIFT_3;
	
	@Column(name = "IOT_TL_1")
	private BigDecimal IOT_TL_1;
	
	@Column(name = "IOT_TL_2")
	private BigDecimal IOT_TL_2;
	
	@Column(name = "IOT_TL_3")
	private BigDecimal IOT_TL_3;
	
	@Column(name = "IOT_TT_1")
	private BigDecimal IOT_TT_1;
	
	@Column(name = "IOT_TT_2")
	private BigDecimal IOT_TT_2;
	
	@Column(name = "IOT_TT_3")
	private BigDecimal IOT_TT_3;
	
	@Column(name = "OFF")
	private BigDecimal OFF;
	
	@Column(name = "SEMI_OFF")
	private BigDecimal SEMI_OFF;
	
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
	
	public WorkDay() {
	}
	
	public WorkDay(WorkDay WorkDay) {
		this.DATE_WD = WorkDay.getDATE_WD();
		this.IWD_SHIFT_1 = WorkDay.getIWD_SHIFT_1();
		this.IWD_SHIFT_2 = WorkDay.getIWD_SHIFT_2();
		this.IWD_SHIFT_3 = WorkDay.getIWD_SHIFT_3();
		this.IOT_TL_1 = WorkDay.getIOT_TL_1();
		this.IOT_TL_2 = WorkDay.getIOT_TL_2();
		this.IOT_TL_3 = WorkDay.getIOT_TL_3();
		this.IOT_TT_1 = WorkDay.getIOT_TT_1();
		this.IOT_TT_2 = WorkDay.getIOT_TT_2();
		this.IOT_TT_3 = WorkDay.getIOT_TT_3();
		this.OFF = WorkDay.getOFF();
		this.SEMI_OFF = WorkDay.getSEMI_OFF();
		this.STATUS = WorkDay.getSTATUS();
		this.CREATION_DATE = WorkDay.getCREATION_DATE();
		this.CREATED_BY = WorkDay.getCREATED_BY();
		this.LAST_UPDATE_DATE = WorkDay.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = WorkDay.getLAST_UPDATED_BY();
	}

	public WorkDay(Date dATE_WD, BigDecimal iWD_SHIFT_1, BigDecimal iWD_SHIFT_2, BigDecimal iWD_SHIFT_3,
			BigDecimal iOT_TL_1, BigDecimal iOT_TL_2, BigDecimal iOT_TL_3, BigDecimal iOT_TT_1, BigDecimal iOT_TT_2,
			BigDecimal iOT_TT_3, BigDecimal oFF, BigDecimal sEMI_OFF, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		DATE_WD = dATE_WD;
		IWD_SHIFT_1 = iWD_SHIFT_1;
		IWD_SHIFT_2 = iWD_SHIFT_2;
		IWD_SHIFT_3 = iWD_SHIFT_3;
		IOT_TL_1 = iOT_TL_1;
		IOT_TL_2 = iOT_TL_2;
		IOT_TL_3 = iOT_TL_3;
		IOT_TT_1 = iOT_TT_1;
		IOT_TT_2 = iOT_TT_2;
		IOT_TT_3 = iOT_TT_3;
		OFF = oFF;
		SEMI_OFF = sEMI_OFF;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public Date getDATE_WD() {
		return DATE_WD;
	}

	public void setDATE_WD(Date dATE_WD) {
		DATE_WD = dATE_WD;
	}

	public BigDecimal getIWD_SHIFT_1() {
		return IWD_SHIFT_1;
	}

	public void setIWD_SHIFT_1(BigDecimal iWD_SHIFT_1) {
		IWD_SHIFT_1 = iWD_SHIFT_1;
	}

	public BigDecimal getIWD_SHIFT_2() {
		return IWD_SHIFT_2;
	}

	public void setIWD_SHIFT_2(BigDecimal iWD_SHIFT_2) {
		IWD_SHIFT_2 = iWD_SHIFT_2;
	}

	public BigDecimal getIWD_SHIFT_3() {
		return IWD_SHIFT_3;
	}

	public void setIWD_SHIFT_3(BigDecimal iWD_SHIFT_3) {
		IWD_SHIFT_3 = iWD_SHIFT_3;
	}

	public BigDecimal getIOT_TL_1() {
		return IOT_TL_1;
	}

	public void setIOT_TL_1(BigDecimal iOT_TL_1) {
		IOT_TL_1 = iOT_TL_1;
	}

	public BigDecimal getIOT_TL_2() {
		return IOT_TL_2;
	}

	public void setIOT_TL_2(BigDecimal iOT_TL_2) {
		IOT_TL_2 = iOT_TL_2;
	}

	public BigDecimal getIOT_TL_3() {
		return IOT_TL_3;
	}

	public void setIOT_TL_3(BigDecimal iOT_TL_3) {
		IOT_TL_3 = iOT_TL_3;
	}

	public BigDecimal getIOT_TT_1() {
		return IOT_TT_1;
	}

	public void setIOT_TT_1(BigDecimal iOT_TT_1) {
		IOT_TT_1 = iOT_TT_1;
	}

	public BigDecimal getIOT_TT_2() {
		return IOT_TT_2;
	}

	public void setIOT_TT_2(BigDecimal iOT_TT_2) {
		IOT_TT_2 = iOT_TT_2;
	}

	public BigDecimal getIOT_TT_3() {
		return IOT_TT_3;
	}

	public void setIOT_TT_3(BigDecimal iOT_TT_3) {
		IOT_TT_3 = iOT_TT_3;
	}

	public BigDecimal getOFF() {
		return OFF;
	}

	public void setOFF(BigDecimal oFF) {
		OFF = oFF;
	}

	public BigDecimal getSEMI_OFF() {
		return SEMI_OFF;
	}

	public void setSEMI_OFF(BigDecimal sEMI_OFF) {
		SEMI_OFF = sEMI_OFF;
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
