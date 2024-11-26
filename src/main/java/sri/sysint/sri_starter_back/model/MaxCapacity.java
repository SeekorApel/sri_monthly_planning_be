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
@Table(name = "SRI_IMPP_D_MAXCAPACITY")
public class MaxCapacity {
	@Id
    @Column(name = "MAX_CAP_ID")
    private BigDecimal MAX_CAP_ID;
	
	@Column(name = "PART_NUMBER")
    private BigDecimal PRODUCT_ID;
	
	@Column(name = "MACHINECURINGTYPE_ID")
	private String MACHINECURINGTYPE_ID;
	
	@Column(name = "CYCLE_TIME")
    private BigDecimal CYCLE_TIME;
	
	@Column(name = "CAPACITY_SHIFT_1")
    private BigDecimal CAPACITY_SHIFT_1;
	
	@Column(name = "CAPACITY_SHIFT_2")
    private BigDecimal CAPACITY_SHIFT_2;
	
	@Column(name = "CAPACITY_SHIFT_3")
    private BigDecimal CAPACITY_SHIFT_3;
	
	
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
	
	public MaxCapacity() {
	}
	
	public MaxCapacity(MaxCapacity maxCapacity) {
		this.MAX_CAP_ID = maxCapacity.getMAX_CAP_ID();
		this.PRODUCT_ID = maxCapacity.getPRODUCT_ID();
		this.MACHINECURINGTYPE_ID = maxCapacity.getMACHINECURINGTYPE_ID();
		this.CYCLE_TIME = maxCapacity.getCYCLE_TIME();
		this.CAPACITY_SHIFT_1 = maxCapacity.getCAPACITY_SHIFT_1();
		this.CAPACITY_SHIFT_2 = maxCapacity.getCAPACITY_SHIFT_2();
		this.CAPACITY_SHIFT_3 = maxCapacity.getCAPACITY_SHIFT_3();
		this.STATUS = maxCapacity.getSTATUS();
		this.CREATION_DATE = maxCapacity.getCREATION_DATE();
		this.CREATED_BY = maxCapacity.getCREATED_BY();
		this.LAST_UPDATE_DATE = maxCapacity.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = maxCapacity.getLAST_UPDATED_BY();
	}

	public MaxCapacity(BigDecimal mAX_CAP_ID, BigDecimal pRODUCT_ID, String mACHINECURINGTYPE_ID, BigDecimal cYCLE_TIME,
			BigDecimal cAPACITY_SHIFT_1, BigDecimal cAPACITY_SHIFT_2, BigDecimal cAPACITY_SHIFT_3, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		MAX_CAP_ID = mAX_CAP_ID;
		PRODUCT_ID = pRODUCT_ID;
		MACHINECURINGTYPE_ID = mACHINECURINGTYPE_ID;
		CYCLE_TIME = cYCLE_TIME;
		CAPACITY_SHIFT_1 = cAPACITY_SHIFT_1;
		CAPACITY_SHIFT_2 = cAPACITY_SHIFT_2;
		CAPACITY_SHIFT_3 = cAPACITY_SHIFT_3;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getMAX_CAP_ID() {
		return MAX_CAP_ID;
	}

	public void setMAX_CAP_ID(BigDecimal mAX_CAP_ID) {
		MAX_CAP_ID = mAX_CAP_ID;
	}

	
	public BigDecimal getPRODUCT_ID() {
		return PRODUCT_ID;
	}

	public void setPRODUCT_ID(BigDecimal pRODUCT_ID) {
		PRODUCT_ID = pRODUCT_ID;
	}

	public String getMACHINECURINGTYPE_ID() {
		return MACHINECURINGTYPE_ID;
	}

	public void setMACHINECURINGTYPE_ID(String mACHINECURINGTYPE_ID) {
		MACHINECURINGTYPE_ID = mACHINECURINGTYPE_ID;
	}

	public BigDecimal getCYCLE_TIME() {
		return CYCLE_TIME;
	}

	public void setCYCLE_TIME(BigDecimal cYCLE_TIME) {
		CYCLE_TIME = cYCLE_TIME;
	}

	public BigDecimal getCAPACITY_SHIFT_1() {
		return CAPACITY_SHIFT_1;
	}

	public void setCAPACITY_SHIFT_1(BigDecimal cAPACITY_SHIFT_1) {
		CAPACITY_SHIFT_1 = cAPACITY_SHIFT_1;
	}

	public BigDecimal getCAPACITY_SHIFT_2() {
		return CAPACITY_SHIFT_2;
	}

	public void setCAPACITY_SHIFT_2(BigDecimal cAPACITY_SHIFT_2) {
		CAPACITY_SHIFT_2 = cAPACITY_SHIFT_2;
	}

	public BigDecimal getCAPACITY_SHIFT_3() {
		return CAPACITY_SHIFT_3;
	}

	public void setCAPACITY_SHIFT_3(BigDecimal cAPACITY_SHIFT_3) {
		CAPACITY_SHIFT_3 = cAPACITY_SHIFT_3;
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
