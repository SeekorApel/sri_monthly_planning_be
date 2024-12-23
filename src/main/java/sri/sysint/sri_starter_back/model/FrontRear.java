package sri.sysint.sri_starter_back.model;

import java.io.Serializable;
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
@Table(name = "SRI_IMPP_M_FRONT_REAR")
public class FrontRear implements Serializable {
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "ID_FRONT_REAR", nullable = false)
    private BigDecimal ID_FRONT_REAR;
	@Column(name = "DETAIL_ID_MO", nullable = false)
	private BigDecimal DETAIL_ID_MO;
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
	
	public FrontRear() {
	}
	
	public FrontRear(FrontRear frontRear) {
		this.ID_FRONT_REAR = frontRear.getID_FRONT_REAR();
		this.DETAIL_ID_MO = frontRear.getDETAIL_ID_MO();
		this.STATUS = frontRear.getSTATUS();
		this.CREATION_DATE = frontRear.getCREATION_DATE();
		this.CREATED_BY = frontRear.getCREATED_BY();
		this.LAST_UPDATE_DATE = frontRear.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = frontRear.getLAST_UPDATED_BY();
	}
	
	

	public FrontRear(BigDecimal iD_FRONT_REAR, BigDecimal dETAIL_ID_MO, BigDecimal sTATUS, Date cREATION_DATE,
			String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		ID_FRONT_REAR = iD_FRONT_REAR;
		DETAIL_ID_MO = dETAIL_ID_MO;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getID_FRONT_REAR() {
		return ID_FRONT_REAR;
	}

	public void setID_FRONT_REAR(BigDecimal iD_FRONT_REAR) {
		ID_FRONT_REAR = iD_FRONT_REAR;
	}

	public BigDecimal getDETAIL_ID_MO() {
		return DETAIL_ID_MO;
	}

	public void setDETAIL_ID_MO(BigDecimal dETAIL_ID_MO) {
		DETAIL_ID_MO = dETAIL_ID_MO;
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
