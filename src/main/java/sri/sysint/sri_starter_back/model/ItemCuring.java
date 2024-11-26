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
@Table(name = "SRI_IMPP_M_ITEMCURING")
public class ItemCuring {
	@Id
    @Column(name = "ITEM_CURING")
    private String ITEM_CURING;
	
	@Column(name = "KAPA_PER_MOULD")
    private BigDecimal KAPA_PER_MOULD;
	
	@Column(name = "NUMBER_OF_MOULD")
    private BigDecimal NUMBER_OF_MOULD;
	
	@Column(name = "MACHINE_TYPE")
    private String MACHINE_TYPE;
	
	@Column(name = "SPARE_MOULD")
    private BigDecimal SPARE_MOULD;
	
	@Column(name = "MOULD_MONTHLY_PLAN")
    private BigDecimal MOULD_MONTHLY_PLAN;
	
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
	
	public ItemCuring() {
	}
	
	public ItemCuring(ItemCuring itemCuring) {
		this.ITEM_CURING = itemCuring.getITEM_CURING();
		this.KAPA_PER_MOULD = itemCuring.getKAPA_PER_MOULD();
		this.NUMBER_OF_MOULD = itemCuring.getNUMBER_OF_MOULD();
		this.MACHINE_TYPE = itemCuring.getMACHINE_TYPE();
		this.SPARE_MOULD = itemCuring.getSPARE_MOULD();
		this.MOULD_MONTHLY_PLAN = itemCuring.getMOULD_MONTHLY_PLAN();
		this.STATUS = itemCuring.getSTATUS();
		this.CREATION_DATE = itemCuring.getCREATION_DATE();
		this.CREATED_BY = itemCuring.getCREATED_BY();
		this.LAST_UPDATE_DATE = itemCuring.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = itemCuring.getLAST_UPDATED_BY();
	}

	public ItemCuring(String iTEM_CURING, BigDecimal kAPA_PER_MOULD, BigDecimal nUMBER_OF_MOULD, String mACHINE_TYPE,
			BigDecimal sPARE_MOULD, BigDecimal mOULD_MONTHLY_PLAN, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		ITEM_CURING = iTEM_CURING;
		KAPA_PER_MOULD = kAPA_PER_MOULD;
		NUMBER_OF_MOULD = nUMBER_OF_MOULD;
		MACHINE_TYPE = mACHINE_TYPE;
		SPARE_MOULD = sPARE_MOULD;
		MOULD_MONTHLY_PLAN = mOULD_MONTHLY_PLAN;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public String getITEM_CURING() {
		return ITEM_CURING;
	}

	public void setITEM_CURING(String iTEM_CURING) {
		ITEM_CURING = iTEM_CURING;
	}

	public BigDecimal getKAPA_PER_MOULD() {
		return KAPA_PER_MOULD;
	}

	public void setKAPA_PER_MOULD(BigDecimal kAPA_PER_MOULD) {
		KAPA_PER_MOULD = kAPA_PER_MOULD;
	}

	public BigDecimal getNUMBER_OF_MOULD() {
		return NUMBER_OF_MOULD;
	}

	public void setNUMBER_OF_MOULD(BigDecimal nUMBER_OF_MOULD) {
		NUMBER_OF_MOULD = nUMBER_OF_MOULD;
	}

	public String getMACHINE_TYPE() {
		return MACHINE_TYPE;
	}

	public void setMACHINE_TYPE(String mACHINE_TYPE) {
		MACHINE_TYPE = mACHINE_TYPE;
	}

	public BigDecimal getSPARE_MOULD() {
		return SPARE_MOULD;
	}

	public void setSPARE_MOULD(BigDecimal sPARE_MOULD) {
		SPARE_MOULD = sPARE_MOULD;
	}


	public BigDecimal getMOULD_MONTHLY_PLAN() {
		return MOULD_MONTHLY_PLAN;
	}

	public void setMOULD_MONTHLY_PLAN(BigDecimal mOULD_MONTHLY_PLAN) {
		MOULD_MONTHLY_PLAN = mOULD_MONTHLY_PLAN;
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
