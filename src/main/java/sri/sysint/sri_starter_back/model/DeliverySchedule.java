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
@Table(name = "SRI_IMPP_M_DELIVERYSCHEDULE")
public class DeliverySchedule {
	@Id
    @Column(name = "DS_ID")
    private BigDecimal DS_ID;
	
	@Column(name = "EFFECTIVE_TIME")
	private Date EFFECTIVE_TIME;
	
	@Column(name = "DATE_ISSUED")
	private Date DATE_ISSUED;
	
	@Column(name = "CATEGORY")
	private String CATEGORY;
	
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
	
	public DeliverySchedule() {
	}
	
	public DeliverySchedule(DeliverySchedule deliverySchedule) {
		
		this.DS_ID = deliverySchedule.getDS_ID();
		this.EFFECTIVE_TIME = deliverySchedule.getEFFECTIVE_TIME();
		this.DATE_ISSUED = deliverySchedule.getDATE_ISSUED();
		this.CATEGORY = deliverySchedule.getCATEGORY();
		this.STATUS = deliverySchedule.getSTATUS();
		this.CREATION_DATE = deliverySchedule.getCREATION_DATE();
		this.CREATED_BY = deliverySchedule.getCREATED_BY();
		this.LAST_UPDATE_DATE = deliverySchedule.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = deliverySchedule.getLAST_UPDATED_BY();
	}
	
	

	public DeliverySchedule(BigDecimal dS_ID, Date eFFECTIVE_TIME, Date dATE_ISSUED, String cATEGORY, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		DS_ID = dS_ID;
		EFFECTIVE_TIME = eFFECTIVE_TIME;
		DATE_ISSUED = dATE_ISSUED;
		CATEGORY = cATEGORY;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getDS_ID() {
		return DS_ID;
	}

	public void setDS_ID(BigDecimal dS_ID) {
		DS_ID = dS_ID;
	}

	public Date getEFFECTIVE_TIME() {
		return EFFECTIVE_TIME;
	}

	public void setEFFECTIVE_TIME(Date eFFECTIVE_TIME) {
		EFFECTIVE_TIME = eFFECTIVE_TIME;
	}

	public Date getDATE_ISSUED() {
		return DATE_ISSUED;
	}

	public void setDATE_ISSUED(Date dATE_ISSUED) {
		DATE_ISSUED = dATE_ISSUED;
	}

	public String getCATEGORY() {
		return CATEGORY;
	}

	public void setCATEGORY(String cATEGORY) {
		CATEGORY = cATEGORY;
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
