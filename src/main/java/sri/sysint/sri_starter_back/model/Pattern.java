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
@Table(name = "SRI_IMPP_M_PATTERN")
public class Pattern {
	@Id
    @Column(name = "PATTERN_ID")
    private BigDecimal PATTERN_ID;
	
	@Column(name = "PATTERN_NAME")
	private String PATTERN_NAME;
	
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
	
	public Pattern() {
	}
	
	public Pattern(Pattern pattern) {
		this.PATTERN_ID = pattern.getPATTERN_ID();
		this.PATTERN_NAME = pattern.getPATTERN_NAME();
		this.STATUS = pattern.getSTATUS();
		this.CREATION_DATE = pattern.getCREATION_DATE();
		this.CREATED_BY = pattern.getCREATED_BY();
		this.LAST_UPDATE_DATE = pattern.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = pattern.getLAST_UPDATED_BY();
	}
	
	public Pattern(BigDecimal pATTERN_ID, String pATTERN_NAME, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		PATTERN_ID = pATTERN_ID;
		PATTERN_NAME = pATTERN_NAME;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getPATTERN_ID() {
		return PATTERN_ID;
	}

	public void setPATTERN_ID(BigDecimal pATTERN_ID) {
		PATTERN_ID = pATTERN_ID;
	}

	public String getPATTERN_NAME() {
		return PATTERN_NAME;
	}

	public void setPATTERN_NAME(String pATTERN_NAME) {
		PATTERN_NAME = pATTERN_NAME;
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
