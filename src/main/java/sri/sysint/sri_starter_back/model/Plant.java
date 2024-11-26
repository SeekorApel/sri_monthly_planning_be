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
@Table(name = "SRI_IMPP_M_PLANT")
public class Plant {
	@Id
    @Column(name = "PLANT_ID")
    private BigDecimal PLANT_ID;
	@Column(name = "PLANT_NAME")
	private String PLANT_NAME;
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
	
	public Plant() {
	}
	
	public Plant(Plant plant) {
		this.PLANT_ID = plant.getPLANT_ID();
		this.PLANT_NAME = plant.getPLANT_NAME();
		this.STATUS = plant.getSTATUS();
		this.CREATION_DATE = plant.getCREATION_DATE();
		this.CREATED_BY = plant.getCREATED_BY();
		this.LAST_UPDATE_DATE = plant.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = plant.getLAST_UPDATED_BY();
	}
	
	public Plant(BigDecimal pLANT_ID, String pLANT_NAME, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		PLANT_ID = pLANT_ID;
		PLANT_NAME = pLANT_NAME;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getPLANT_ID() {
		return PLANT_ID;
	}

	public void setPLANT_ID(BigDecimal pLANT_ID) {
		PLANT_ID = pLANT_ID;
	}

	public String getPLANT_NAME() {
		return PLANT_NAME;
	}

	public void setPLANT_NAME(String pLANT_NAME) {
		PLANT_NAME = pLANT_NAME;
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
