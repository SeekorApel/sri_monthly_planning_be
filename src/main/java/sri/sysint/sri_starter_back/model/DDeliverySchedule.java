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
@Table(name = "SRI_IMPP_D_DELIVERYSCHEDULE")
public class DDeliverySchedule {
	@Id
    @Column(name = "DETAIL_DS_ID")
    private BigDecimal DETAIL_DS_ID;
	
    @Column(name = "DS_ID")
    private BigDecimal DS_ID;
    
	@Column(name = "PART_NUM")
	private BigDecimal PART_NUM;
	
	@Column(name = "DATE_DS")
	private Date DATE_DS;
	
    @Column(name = "TOTAL_DELIVERY")
    private BigDecimal TOTAL_DELIVERY;
	
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
	
	public DDeliverySchedule() {
	}
	
	public DDeliverySchedule(DDeliverySchedule dDeliverySchedule) {
		this.DETAIL_DS_ID = dDeliverySchedule.getDETAIL_DS_ID();
		this.DS_ID = dDeliverySchedule.getDS_ID();
		this.PART_NUM = dDeliverySchedule.getPART_NUM();
		this.DATE_DS = dDeliverySchedule.getDATE_DS();
		this.TOTAL_DELIVERY = dDeliverySchedule.getTOTAL_DELIVERY();
		this.STATUS = dDeliverySchedule.getSTATUS();
		this.CREATION_DATE = dDeliverySchedule.getCREATION_DATE();
		this.CREATED_BY = dDeliverySchedule.getCREATED_BY();
		this.LAST_UPDATE_DATE = dDeliverySchedule.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = dDeliverySchedule.getLAST_UPDATED_BY();
	}
	
	
	public DDeliverySchedule(BigDecimal dETAIL_DS_ID, BigDecimal dS_ID, BigDecimal pART_NUM, Date dATE_DS,
			BigDecimal tOTAL_DELIVERY, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE,
			String lAST_UPDATED_BY) {
		super();
		DETAIL_DS_ID = dETAIL_DS_ID;
		DS_ID = dS_ID;
		PART_NUM = pART_NUM;
		DATE_DS = dATE_DS;
		TOTAL_DELIVERY = tOTAL_DELIVERY;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getDETAIL_DS_ID() {
		return DETAIL_DS_ID;
	}

	public void setDETAIL_DS_ID(BigDecimal dETAIL_DS_ID) {
		DETAIL_DS_ID = dETAIL_DS_ID;
	}

	public BigDecimal getDS_ID() {
		return DS_ID;
	}

	public void setDS_ID(BigDecimal dS_ID) {
		DS_ID = dS_ID;
	}
	
	public BigDecimal getPART_NUM() {
		return PART_NUM;
	}

	public void setPART_NUM(BigDecimal pART_NUM) {
		PART_NUM = pART_NUM;
	}

	public Date getDATE_DS() {
		return DATE_DS;
	}

	public void setDATE_DS(Date dATE_DS) {
		DATE_DS = dATE_DS;
	}

	public BigDecimal getTOTAL_DELIVERY() {
		return TOTAL_DELIVERY;
	}

	public void setTOTAL_DELIVERY(BigDecimal tOTAL_DELIVERY) {
		TOTAL_DELIVERY = tOTAL_DELIVERY;
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
