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
@Table(name = "SRI_IMPP_D_QUADRANTDISTANCE")
public class QuadrantDistance {
	@Id
    @Column(name = "ID_Q_DISTANCE")
    private BigDecimal ID_Q_DISTANCE;
	
    @Column(name = "QUADRANT_ID_1")
    private BigDecimal QUADRANT_ID_1;
    
    @Column(name = "QUADRANT_ID_2")
    private BigDecimal QUADRANT_ID_2;
    
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
	
	public QuadrantDistance() {
	}
	
	public QuadrantDistance(QuadrantDistance quadrantDistance) {
		this.ID_Q_DISTANCE = quadrantDistance.getID_Q_DISTANCE();
		this.QUADRANT_ID_1 = quadrantDistance.getQUADRANT_ID_1();
		this.QUADRANT_ID_2 = quadrantDistance.getQUADRANT_ID_2();
		this.DISTANCE = quadrantDistance.getDISTANCE();
		this.STATUS = quadrantDistance.getSTATUS();
		this.CREATION_DATE = quadrantDistance.getCREATION_DATE();
		this.CREATED_BY = quadrantDistance.getCREATED_BY();
		this.LAST_UPDATE_DATE = quadrantDistance.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = quadrantDistance.getLAST_UPDATED_BY();
	}
	
	

	public QuadrantDistance(BigDecimal iD_Q_DISTANCE, BigDecimal qUADRANT_ID_1, BigDecimal qUADRANT_ID_2,
			BigDecimal dISTANCE, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE,
			String lAST_UPDATED_BY) {
		super();
		ID_Q_DISTANCE = iD_Q_DISTANCE;
		QUADRANT_ID_1 = qUADRANT_ID_1;
		QUADRANT_ID_2 = qUADRANT_ID_2;
		DISTANCE = dISTANCE;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getID_Q_DISTANCE() {
		return ID_Q_DISTANCE;
	}

	public void setID_Q_DISTANCE(BigDecimal iD_Q_DISTANCE) {
		ID_Q_DISTANCE = iD_Q_DISTANCE;
	}

	public BigDecimal getQUADRANT_ID_1() {
		return QUADRANT_ID_1;
	}

	public void setQUADRANT_ID_1(BigDecimal qUADRANT_ID_1) {
		QUADRANT_ID_1 = qUADRANT_ID_1;
	}

	public BigDecimal getQUADRANT_ID_2() {
		return QUADRANT_ID_2;
	}

	public void setQUADRANT_ID_2(BigDecimal qUADRANT_ID_2) {
		QUADRANT_ID_2 = qUADRANT_ID_2;
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
