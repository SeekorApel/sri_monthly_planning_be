package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_D_TASSSIZE")
public class TassSize {
	
	@Id
    @Column(name = "TASSIZE_ID")
    private BigDecimal TASSIZE_ID;
	
    @Column(name = "MACHINETASSTYPE_ID")
    private String MACHINETASSTYPE_ID;
	
	@Column(name = "SIZE_ID")
	private String SIZE_ID;
	
	@Column(name = "CAPACITY")
	private BigDecimal CAPACITY;
	
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
	
	public TassSize() {}
	
	public TassSize(TassSize tassSize) {
		this.TASSIZE_ID = tassSize.getTASSIZE_ID();
		this.MACHINETASSTYPE_ID = tassSize.getMACHINETASSTYPE_ID();
		this.SIZE_ID = tassSize.getSIZE_ID();
		this.CAPACITY = tassSize.getCAPACITY();
		this.STATUS = tassSize.getSTATUS();
		this.CREATION_DATE = tassSize.getCREATION_DATE();
		this.CREATED_BY = tassSize.getCREATED_BY();
		this.LAST_UPDATE_DATE = tassSize.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = tassSize.getLAST_UPDATED_BY();
	}

	public TassSize(BigDecimal tASS_SIZE, String mACHINETASSTYPE_ID, String sIZE_ID, BigDecimal cAPACITY, BigDecimal sTATUS,
			Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		TASSIZE_ID = tASS_SIZE;
		MACHINETASSTYPE_ID = mACHINETASSTYPE_ID;
		SIZE_ID = sIZE_ID;
		CAPACITY = cAPACITY;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}


	

	public BigDecimal getTASSIZE_ID() {
		return TASSIZE_ID;
	}

	public void setTASSIZE_ID(BigDecimal tASSIZE_ID) {
		TASSIZE_ID = tASSIZE_ID;
	}

	public String getMACHINETASSTYPE_ID() {
		return MACHINETASSTYPE_ID;
	}

	public void setMACHINETASSTYPE_ID(String mACHINETASSTYPE_ID) {
		MACHINETASSTYPE_ID = mACHINETASSTYPE_ID;
	}

	public String getSIZE_ID() {
		return SIZE_ID;
	}

	public void setSIZE_ID(String sIZE_ID) {
		SIZE_ID = sIZE_ID;
	}

	public BigDecimal getCAPACITY() {
		return CAPACITY;
	}

	public void setCAPACITY(BigDecimal cAPACITY) {
		CAPACITY = cAPACITY;
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
