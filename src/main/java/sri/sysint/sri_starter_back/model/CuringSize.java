package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_D_CURINGSIZE")
public class CuringSize {
	@Id
    @Column(name = "CURINGSIZE_ID")
    private BigDecimal CURINGSIZE_ID;
	
    @Column(name = "MACHINECURINGTYPE_ID")
    private String MACHINECURINGTYPE_ID;
	
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
	
	public CuringSize() {}
	
	public CuringSize(CuringSize curingSize) {
		this.CURINGSIZE_ID = curingSize.getCURINGSIZE_ID();
		this.MACHINECURINGTYPE_ID = curingSize.getMACHINECURINGTYPE_ID();
		this.SIZE_ID = curingSize.getSIZE_ID();
		this.CAPACITY = curingSize.getCAPACITY();
		this.STATUS = curingSize.getSTATUS();
		this.CREATION_DATE = curingSize.getCREATION_DATE();
		this.CREATED_BY = curingSize.getCREATED_BY();
		this.LAST_UPDATE_DATE = curingSize.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = curingSize.getLAST_UPDATED_BY();
	}


	public CuringSize(BigDecimal cURINGSIZE_ID, String mACHINECURINGTYPE_ID, String sIZE_ID, BigDecimal cAPACITY,
			BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		CURINGSIZE_ID = cURINGSIZE_ID;
		MACHINECURINGTYPE_ID = mACHINECURINGTYPE_ID;
		SIZE_ID = sIZE_ID;
		CAPACITY = cAPACITY;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getCURINGSIZE_ID() {
		return CURINGSIZE_ID;
	}

	public void setCURINGSIZE_ID(BigDecimal cURINGSIZE_ID) {
		CURINGSIZE_ID = cURINGSIZE_ID;
	}

	public String getMACHINECURINGTYPE_ID() {
		return MACHINECURINGTYPE_ID;
	}

	public void setMACHINECURINGTYPE_ID(String mACHINECURINGTYPE_ID) {
		MACHINECURINGTYPE_ID = mACHINECURINGTYPE_ID;
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
