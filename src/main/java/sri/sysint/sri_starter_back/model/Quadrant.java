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
@Table(name = "SRI_IMPP_M_QUADRANT")
public class Quadrant {
	@Id
    @Column(name = "QUADRANT_ID")
    private BigDecimal QUADRANT_ID;
	
	@Column(name = "BUILDING_ID")
    private BigDecimal BUILDING_ID;
	
	@Column(name = "QUADRANT_NAME")
	private String QUADRANT_NAME;
	
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
	
	public Quadrant() {
	}
	
	public Quadrant(Quadrant quadrant) {
		this.QUADRANT_ID = quadrant.getQUADRANT_ID();
		this.BUILDING_ID = quadrant.getBUILDING_ID();
		this.QUADRANT_NAME = quadrant.getQUADRANT_NAME();
		this.STATUS = quadrant.getSTATUS();
		this.CREATION_DATE = quadrant.getCREATION_DATE();
		this.CREATED_BY = quadrant.getCREATED_BY();
		this.LAST_UPDATE_DATE = quadrant.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = quadrant.getLAST_UPDATED_BY();
	}

	public Quadrant(BigDecimal qUADRANT_ID, BigDecimal bUILDING_ID, String qUADRANT_NAME, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		QUADRANT_ID = qUADRANT_ID;
		BUILDING_ID = bUILDING_ID;
		QUADRANT_NAME = qUADRANT_NAME;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getQUADRANT_ID() {
		return QUADRANT_ID;
	}

	public void setQUADRANT_ID(BigDecimal qUADRANT_ID) {
		QUADRANT_ID = qUADRANT_ID;
	}

	public BigDecimal getBUILDING_ID() {
		return BUILDING_ID;
	}

	public void setBUILDING_ID(BigDecimal bUILDING_ID) {
		BUILDING_ID = bUILDING_ID;
	}

	public String getQUADRANT_NAME() {
		return QUADRANT_NAME;
	}

	public void setQUADRANT_NAME(String qUADRANT_NAME) {
		QUADRANT_NAME = qUADRANT_NAME;
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
