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
@Table(name = "SRI_IMPP_D_BUILDINGDISTANCE")
public class BuildingDistance {
	@Id
    @Column(name = "ID_B_DISTANCE")
    private BigDecimal ID_B_DISTANCE;
	
    @Column(name = "BUILDING_ID_1")
    private BigDecimal BUILDING_ID_1;
    
    @Column(name = "BUILDING_ID_2")
    private BigDecimal BUILDING_ID_2;
    
    @Column(name = "DISTANCE")
    private BigDecimal DISTANCE;
   
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
	
	public BuildingDistance() {
	}
	
	public BuildingDistance(BuildingDistance buildingDistance) {
		this.ID_B_DISTANCE = buildingDistance.getID_B_DISTANCE();
		this.BUILDING_ID_1 = buildingDistance.getBUILDING_ID_1();
		this.BUILDING_ID_2 = buildingDistance.getBUILDING_ID_2();
		this.DISTANCE = buildingDistance.getDISTANCE();
		this.STATUS = buildingDistance.getSTATUS();
		this.CREATION_DATE = buildingDistance.getCREATION_DATE();
		this.CREATED_BY = buildingDistance.getCREATED_BY();
		this.LAST_UPDATE_DATE = buildingDistance.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = buildingDistance.getLAST_UPDATED_BY();
	}

	public BuildingDistance(BigDecimal iD_B_DISTANCE, BigDecimal bUILDING_ID_1, BigDecimal bUILDING_ID_2,
			BigDecimal dISTANCE, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE,
			String lAST_UPDATED_BY) {
		super();
		ID_B_DISTANCE = iD_B_DISTANCE;
		BUILDING_ID_1 = bUILDING_ID_1;
		BUILDING_ID_2 = bUILDING_ID_2;
		DISTANCE = dISTANCE;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getID_B_DISTANCE() {
		return ID_B_DISTANCE;
	}

	public void setID_B_DISTANCE(BigDecimal iD_B_DISTANCE) {
		ID_B_DISTANCE = iD_B_DISTANCE;
	}

	public BigDecimal getBUILDING_ID_1() {
		return BUILDING_ID_1;
	}

	public void setBUILDING_ID_1(BigDecimal bUILDING_ID_1) {
		BUILDING_ID_1 = bUILDING_ID_1;
	}

	public BigDecimal getBUILDING_ID_2() {
		return BUILDING_ID_2;
	}

	public void setBUILDING_ID_2(BigDecimal bUILDING_ID_2) {
		BUILDING_ID_2 = bUILDING_ID_2;
	}

	public BigDecimal getDISTANCE() {
		return DISTANCE;
	}

	public void setDISTANCE(BigDecimal dISTANCE) {
		DISTANCE = dISTANCE;
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
