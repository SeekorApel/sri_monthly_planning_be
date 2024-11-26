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
@Table(name = "SRI_IMPP_M_BUILDING")
public class Building {
	@Id
    @Column(name = "BUILDING_ID")
    private BigDecimal BUILDING_ID;
	
	@Column(name = "PLANT_ID")
    private BigDecimal PLANT_ID;
	
	@Column(name = "BUILDING_NAME")
	private String BUILDING_NAME;
	
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
	
	public Building() {
	}
	
	public Building(Building building) {
		this.BUILDING_ID = building.getBUILDING_ID();
		this.PLANT_ID = building.getPLANT_ID();
		this.BUILDING_NAME = building.getBUILDING_NAME();
		this.STATUS = building.getSTATUS();
		this.CREATION_DATE = building.getCREATION_DATE();
		this.CREATED_BY = building.getCREATED_BY();
		this.LAST_UPDATE_DATE = building.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = building.getLAST_UPDATED_BY();
	}

	public Building(BigDecimal bUILDING_ID, BigDecimal pLANT_ID, String bUILDING_NAME, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		BUILDING_ID = bUILDING_ID;
		PLANT_ID = pLANT_ID;
		BUILDING_NAME = bUILDING_NAME;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getBUILDING_ID() {
		return BUILDING_ID;
	}

	public void setBUILDING_ID(BigDecimal bUILDING_ID) {
		BUILDING_ID = bUILDING_ID;
	}

	public BigDecimal getPLANT_ID() {
		return PLANT_ID;
	}

	public void setPLANT_ID(BigDecimal pLANT_ID) {
		PLANT_ID = pLANT_ID;
	}

	public String getBUILDING_NAME() {
		return BUILDING_NAME;
	}

	public void setBUILDING_NAME(String bUILDING_NAME) {
		BUILDING_NAME = bUILDING_NAME;
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
