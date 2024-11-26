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
public class Kapa {
	@Id
    @Column(name = "PART_NUMBER")
    private BigDecimal PART_NUMBER;
	
	@Column(name = "ITEM_CURING")
	private String ITEM_CURING;
	
	@Column(name = "TYPE_CURING")
	private String TYPE_CURING;
	
	@Column(name = "DESCRIPTION")
	private String DESCRIPTION;
	
    @Column(name = "CYCLE_TIME")
    private BigDecimal CYCLE_TIME;
    
    @Column(name = "SHIFT")
    private BigDecimal SHIFT;
    
    @Column(name = "KAPA_PERSHIFT")
    private BigDecimal KAPA_PERSHIFT;
    
    @Column(name = "LAST_UPDATE_DATA")
    private BigDecimal LAST_UPDATE_DATA;
    
	@Column(name = "MACHINE")
	private String MACHINE;
	
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
	
	public Kapa() {
	}
	
	public Kapa(Kapa kapa) {
		this.PART_NUMBER = kapa.getPART_NUMBER();
		this.ITEM_CURING = kapa.getITEM_CURING();
		this.TYPE_CURING = kapa.getTYPE_CURING();
		this.DESCRIPTION = kapa.getDESCRIPTION();
		this.CYCLE_TIME = kapa.getCYCLE_TIME();
		this.SHIFT = kapa.getSHIFT();
		this.KAPA_PERSHIFT = kapa.getKAPA_PERSHIFT();
		this.LAST_UPDATE_DATA = kapa.getLAST_UPDATE_DATA();
		this.MACHINE = kapa.getMACHINE();
		this.STATUS = kapa.getSTATUS();
		this.CREATION_DATE = kapa.getCREATION_DATE();
		this.CREATED_BY = kapa.getCREATED_BY();
		this.LAST_UPDATE_DATE = kapa.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = kapa.getLAST_UPDATED_BY();
	}
	
	

	public Kapa(BigDecimal pART_NUMBER, String iTEM_CURING, String tYPE_CURING, String dESCRIPTION,
			BigDecimal cYCLE_TIME, BigDecimal sHIFT, BigDecimal kAPA_PERSHIFT, BigDecimal lAST_UPDATE_DATA,
			String mACHINE, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE,
			String lAST_UPDATED_BY) {
		super();
		PART_NUMBER = pART_NUMBER;
		ITEM_CURING = iTEM_CURING;
		TYPE_CURING = tYPE_CURING;
		DESCRIPTION = dESCRIPTION;
		CYCLE_TIME = cYCLE_TIME;
		SHIFT = sHIFT;
		KAPA_PERSHIFT = kAPA_PERSHIFT;
		LAST_UPDATE_DATA = lAST_UPDATE_DATA;
		MACHINE = mACHINE;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getPART_NUMBER() {
		return PART_NUMBER;
	}

	public void setPART_NUMBER(BigDecimal pART_NUMBER) {
		PART_NUMBER = pART_NUMBER;
	}

	public String getITEM_CURING() {
		return ITEM_CURING;
	}

	public void setITEM_CURING(String iTEM_CURING) {
		ITEM_CURING = iTEM_CURING;
	}

	public String getTYPE_CURING() {
		return TYPE_CURING;
	}

	public void setTYPE_CURING(String tYPE_CURING) {
		TYPE_CURING = tYPE_CURING;
	}

	public String getDESCRIPTION() {
		return DESCRIPTION;
	}

	public void setDESCRIPTION(String dESCRIPTION) {
		DESCRIPTION = dESCRIPTION;
	}

	public BigDecimal getCYCLE_TIME() {
		return CYCLE_TIME;
	}

	public void setCYCLE_TIME(BigDecimal cYCLE_TIME) {
		CYCLE_TIME = cYCLE_TIME;
	}

	public BigDecimal getSHIFT() {
		return SHIFT;
	}

	public void setSHIFT(BigDecimal sHIFT) {
		SHIFT = sHIFT;
	}

	public BigDecimal getKAPA_PERSHIFT() {
		return KAPA_PERSHIFT;
	}

	public void setKAPA_PERSHIFT(BigDecimal kAPA_PERSHIFT) {
		KAPA_PERSHIFT = kAPA_PERSHIFT;
	}

	public BigDecimal getLAST_UPDATE_DATA() {
		return LAST_UPDATE_DATA;
	}

	public void setLAST_UPDATE_DATA(BigDecimal lAST_UPDATE_DATA) {
		LAST_UPDATE_DATA = lAST_UPDATE_DATA;
	}

	public String getMACHINE() {
		return MACHINE;
	}

	public void setMACHINE(String mACHINE) {
		MACHINE = mACHINE;
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
