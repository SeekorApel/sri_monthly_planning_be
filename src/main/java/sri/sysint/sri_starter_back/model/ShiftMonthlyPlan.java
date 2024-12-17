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
@Table(name = "SRI_IMPP_D_SHIFTMONTHLYPLAN_CURING")
public class ShiftMonthlyPlan {
	@Id
    @Column(name = "DETAIL_SHIFT_ID_CURING")
    private BigDecimal DETAIL_SHIFT_ID_CURING;
	
	@Column(name = "DETAIL_DAILY_ID_CURING")
    private BigDecimal DETAIL_DAILY_ID_CURING;

	@Column(name = "MO_ID")
	private String MO_ID;
	
	@Column(name = "MP_CURING_ID")
	private String MP_CURING_ID;

	@Column(name = "ACTUAL_DATE")
	private Date DATE;
	
	@Column(name = "ITEM_CURING")
	private String ITEM_CURING;
	
	@Column(name = "PART_NUMBER")
	private BigDecimal PART_NUMBER;
	
	@Column(name = "WORK_CENTER_TEXT")
	private String WORK_CENTER_TEXT;
	
	@Column(name = "CAVITY")
    private BigDecimal CAVITY;
	
	@Column(name = "CAVITY_USAGE")
    private BigDecimal CAVITY_USAGE;
	
	@Column(name = "WH_SHIFT_1")
    private BigDecimal WH_SHIFT_1;
	
	@Column(name = "WH_SHIFT_2")
    private BigDecimal WH_SHIFT_2;
	
	@Column(name = "WH_SHIFT_3")
    private BigDecimal WH_SHIFT_3;
	
	@Column(name = "KAPA_SHIFT_1")
    private BigDecimal KAPA_SHIFT_1;
	
	@Column(name = "KAPA_SHIFT_2")
    private BigDecimal KAPA_SHIFT_2;
	
	@Column(name = "KAPA_SHIFT_3")
    private BigDecimal KAPA_SHIFT_3;
	
	@Column(name = "TOTAL_KAPA")
    private BigDecimal TOTAL_KAPA;
	
	@Column(name = "CAVITY_EXIST")
    private BigDecimal CAVITY_EXIST;
	
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
	
	public ShiftMonthlyPlan() {
	}
	
	public ShiftMonthlyPlan(BigDecimal DETAIL_SHIFT_ID_CURING, BigDecimal DETAIL_DAILY_ID_CURING, String MO_ID, 
            String MP_CURING_ID, Date DATE, String ITEM_CURING, BigDecimal PART_NUMBER, 
            String WORK_CENTER_TEXT, BigDecimal CAVITY, BigDecimal CAVITY_USAGE, 
            BigDecimal WH_SHIFT_1, BigDecimal WH_SHIFT_2, BigDecimal WH_SHIFT_3, 
            BigDecimal KAPA_SHIFT_1, BigDecimal KAPA_SHIFT_2, BigDecimal KAPA_SHIFT_3, 
            BigDecimal TOTAL_KAPA, BigDecimal CAVITY_EXIST, BigDecimal STATUS, 
            Date CREATION_DATE, String CREATED_BY, Date LAST_UPDATE_DATE, 
            String LAST_UPDATED_BY) {
		this.DETAIL_SHIFT_ID_CURING = DETAIL_SHIFT_ID_CURING;
		this.DETAIL_DAILY_ID_CURING = DETAIL_DAILY_ID_CURING;
		this.MO_ID = MO_ID;
		this.MP_CURING_ID = MP_CURING_ID;
		this.DATE = DATE;
		this.ITEM_CURING = ITEM_CURING;
		this.PART_NUMBER = PART_NUMBER;
		this.WORK_CENTER_TEXT = WORK_CENTER_TEXT;
		this.CAVITY = CAVITY;
		this.CAVITY_USAGE = CAVITY_USAGE;
		this.WH_SHIFT_1 = WH_SHIFT_1;
		this.WH_SHIFT_2 = WH_SHIFT_2;
		this.WH_SHIFT_3 = WH_SHIFT_3;
		this.KAPA_SHIFT_1 = KAPA_SHIFT_1;
		this.KAPA_SHIFT_2 = KAPA_SHIFT_2;
		this.KAPA_SHIFT_3 = KAPA_SHIFT_3;
		this.TOTAL_KAPA = TOTAL_KAPA;
		this.CAVITY_EXIST = CAVITY_EXIST;
		this.STATUS = STATUS;
		this.CREATION_DATE = CREATION_DATE;
		this.CREATED_BY = CREATED_BY;
		this.LAST_UPDATE_DATE = LAST_UPDATE_DATE;
		this.LAST_UPDATED_BY = LAST_UPDATED_BY;
	}

	public ShiftMonthlyPlan(ShiftMonthlyPlan existingPlan) {
	    this.DETAIL_SHIFT_ID_CURING = existingPlan.DETAIL_SHIFT_ID_CURING;
	    this.DETAIL_DAILY_ID_CURING = existingPlan.DETAIL_DAILY_ID_CURING;
	    this.MO_ID = existingPlan.MO_ID;
	    this.MP_CURING_ID = existingPlan.MP_CURING_ID;
	    this.DATE = existingPlan.DATE;
	    this.ITEM_CURING = existingPlan.ITEM_CURING;
	    this.PART_NUMBER = existingPlan.PART_NUMBER;
	    this.WORK_CENTER_TEXT = existingPlan.WORK_CENTER_TEXT;
	    this.CAVITY = existingPlan.CAVITY;
	    this.CAVITY_USAGE = existingPlan.CAVITY_USAGE;
	    this.WH_SHIFT_1 = existingPlan.WH_SHIFT_1;
	    this.WH_SHIFT_2 = existingPlan.WH_SHIFT_2;
	    this.WH_SHIFT_3 = existingPlan.WH_SHIFT_3;
	    this.KAPA_SHIFT_1 = existingPlan.KAPA_SHIFT_1;
	    this.KAPA_SHIFT_2 = existingPlan.KAPA_SHIFT_2;
	    this.KAPA_SHIFT_3 = existingPlan.KAPA_SHIFT_3;
	    this.TOTAL_KAPA = existingPlan.TOTAL_KAPA;
	    this.CAVITY_EXIST = existingPlan.CAVITY_EXIST;
	    this.STATUS = existingPlan.STATUS;
	    this.CREATION_DATE = existingPlan.CREATION_DATE;
	    this.CREATED_BY = existingPlan.CREATED_BY;
	    this.LAST_UPDATE_DATE = existingPlan.LAST_UPDATE_DATE;
	    this.LAST_UPDATED_BY = existingPlan.LAST_UPDATED_BY;
	}

	public BigDecimal getDETAIL_SHIFT_ID_CURING() {
		return DETAIL_SHIFT_ID_CURING;
	}

	public void setDETAIL_SHIFT_ID_CURING(BigDecimal dETAIL_SHIFT_ID_CURING) {
		DETAIL_SHIFT_ID_CURING = dETAIL_SHIFT_ID_CURING;
	}

	public BigDecimal getDETAIL_DAILY_ID_CURING() {
		return DETAIL_DAILY_ID_CURING;
	}

	public void setDETAIL_DAILY_ID_CURING(BigDecimal dETAIL_DAILY_ID_CURING) {
		DETAIL_DAILY_ID_CURING = dETAIL_DAILY_ID_CURING;
	}

	public String getMO_ID() {
		return MO_ID;
	}

	public void setMO_ID(String mO_ID) {
		MO_ID = mO_ID;
	}

	public String getMP_CURING_ID() {
		return MP_CURING_ID;
	}

	public void setMP_CURING_ID(String mP_CURING_ID) {
		MP_CURING_ID = mP_CURING_ID;
	}

	public Date getDATE() {
		return DATE;
	}

	public void setDATE(Date dATE) {
		DATE = dATE;
	}

	public String getITEM_CURING() {
		return ITEM_CURING;
	}

	public void setITEM_CURING(String iTEM_CURING) {
		ITEM_CURING = iTEM_CURING;
	}

	public BigDecimal getPART_NUMBER() {
		return PART_NUMBER;
	}

	public void setPART_NUMBER(BigDecimal pART_NUMBER) {
		PART_NUMBER = pART_NUMBER;
	}

	public String getWORK_CENTER_TEXT() {
		return WORK_CENTER_TEXT;
	}

	public void setWORK_CENTER_TEXT(String wORK_CENTER_TEXT) {
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
	}

	public BigDecimal getCAVITY() {
		return CAVITY;
	}

	public void setCAVITY(BigDecimal cAVITY) {
		CAVITY = cAVITY;
	}

	public BigDecimal getCAVITY_USAGE() {
		return CAVITY_USAGE;
	}

	public void setCAVITY_USAGE(BigDecimal cAVITY_USAGE) {
		CAVITY_USAGE = cAVITY_USAGE;
	}

	public BigDecimal getWH_SHIFT_1() {
		return WH_SHIFT_1;
	}

	public void setWH_SHIFT_1(BigDecimal wH_SHIFT_1) {
		WH_SHIFT_1 = wH_SHIFT_1;
	}

	public BigDecimal getWH_SHIFT_2() {
		return WH_SHIFT_2;
	}

	public void setWH_SHIFT_2(BigDecimal wH_SHIFT_2) {
		WH_SHIFT_2 = wH_SHIFT_2;
	}

	public BigDecimal getWH_SHIFT_3() {
		return WH_SHIFT_3;
	}

	public void setWH_SHIFT_3(BigDecimal wH_SHIFT_3) {
		WH_SHIFT_3 = wH_SHIFT_3;
	}

	public BigDecimal getKAPA_SHIFT_1() {
		return KAPA_SHIFT_1;
	}

	public void setKAPA_SHIFT_1(BigDecimal kAPA_SHIFT_1) {
		KAPA_SHIFT_1 = kAPA_SHIFT_1;
	}

	public BigDecimal getKAPA_SHIFT_2() {
		return KAPA_SHIFT_2;
	}

	public void setKAPA_SHIFT_2(BigDecimal kAPA_SHIFT_2) {
		KAPA_SHIFT_2 = kAPA_SHIFT_2;
	}

	public BigDecimal getKAPA_SHIFT_3() {
		return KAPA_SHIFT_3;
	}

	public void setKAPA_SHIFT_3(BigDecimal kAPA_SHIFT_3) {
		KAPA_SHIFT_3 = kAPA_SHIFT_3;
	}

	public BigDecimal getTOTAL_KAPA() {
		return TOTAL_KAPA;
	}

	public void setTOTAL_KAPA(BigDecimal tOTAL_KAPA) {
		TOTAL_KAPA = tOTAL_KAPA;
	}

	public BigDecimal getCAVITY_EXIST() {
		return CAVITY_EXIST;
	}

	public void setCAVITY_EXIST(BigDecimal cAVITY_EXIST) {
		CAVITY_EXIST = cAVITY_EXIST;
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
