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
@Table(name = "SRI_IMPP_D_WD")
public class DWorkDay {
	@Id
    @Column(name = "DETAIL_WD_ID")
    private BigDecimal DETAIL_WD_ID;
	
	@Column(name = "DATE_WD")
	private Date DATE_WD;
	
	@Column(name = "PARENT")
	private String PARENT;
	
	@Column(name = "DESCRIPTION")
	private String DESCRIPTION;
	
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
	
	public DWorkDay() {
	}
	
	public DWorkDay(DWorkDay dWorkDay) {
		this.DETAIL_WD_ID = dWorkDay.getDETAIL_WD_ID();
		this.DATE_WD = dWorkDay.getDATE_WD();
		this.PARENT = dWorkDay.getPARENT();
		this.DESCRIPTION = dWorkDay.getDESCRIPTION();
		this.STATUS = dWorkDay.getSTATUS();
		this.CREATION_DATE = dWorkDay.getCREATION_DATE();
		this.CREATED_BY = dWorkDay.getCREATED_BY();
		this.LAST_UPDATE_DATE = dWorkDay.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = dWorkDay.getLAST_UPDATED_BY();
	}

	public DWorkDay(BigDecimal dETAIL_WD_ID, Date dATE_WD, String pARENT, String dESCRIPTION, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		DETAIL_WD_ID = dETAIL_WD_ID;
		DATE_WD = dATE_WD;
		PARENT = pARENT;
		DESCRIPTION = dESCRIPTION;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getDETAIL_WD_ID() {
		return DETAIL_WD_ID;
	}

	public void setDETAIL_WD_ID(BigDecimal dETAIL_WD_ID) {
		DETAIL_WD_ID = dETAIL_WD_ID;
	}

	public Date getDATE_WD() {
		return DATE_WD;
	}

	public void setDATE_WD(Date dATE_WD) {
		DATE_WD = dATE_WD;
	}

	public String getPARENT() {
		return PARENT;
	}

	public void setPARENT(String pARENT) {
		PARENT = pARENT;
	}

	public String getDESCRIPTION() {
		return DESCRIPTION;
	}

	public void setDESCRIPTION(String dESCRIPTION) {
		DESCRIPTION = dESCRIPTION;
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
