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
@Table(name = "SRI_IMPP_M_PRODUCTTYPE")
public class ProductType {
	@Id
    @Column(name = "PRODUCT_TYPE_ID")
    private BigDecimal PRODUCT_TYPE_ID;
	
	@Column(name = "PRODUCT_MERK")
	private String PRODUCT_MERK;
	
	@Column(name = "PRODUCT_TYPE")
	private String PRODUCT_TYPE;
	
	@Column(name = "CATEGORY")
	private String CATEGORY;
	
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
	
	public ProductType() {
		
	}
	public ProductType(ProductType productType) {
		this.PRODUCT_TYPE_ID = productType.getPRODUCT_TYPE_ID();
		this.PRODUCT_MERK = productType.getPRODUCT_MERK();
		this.PRODUCT_TYPE = productType.getPRODUCT_TYPE();
		this.CATEGORY = productType.getCATEGORY();
		this.STATUS = productType.getSTATUS();
		this.CREATION_DATE = productType.getCREATION_DATE();
		this.CREATED_BY = productType.getCREATED_BY();
		this.LAST_UPDATE_DATE = productType.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = productType.getLAST_UPDATED_BY();
	}

	public ProductType(BigDecimal pRODUCT_TYPE_ID, String pRODUCT_MERK, String pRODUCT_TYPE, String cATEGORY,
			BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		PRODUCT_TYPE_ID = pRODUCT_TYPE_ID;
		PRODUCT_MERK = pRODUCT_MERK;
		PRODUCT_TYPE = pRODUCT_TYPE;
		CATEGORY = cATEGORY;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}


	public BigDecimal getPRODUCT_TYPE_ID() {
		return PRODUCT_TYPE_ID;
	}

	public void setPRODUCT_TYPE_ID(BigDecimal pRODUCT_TYPE_ID) {
		PRODUCT_TYPE_ID = pRODUCT_TYPE_ID;
	}

	public String getPRODUCT_MERK() {
		return PRODUCT_MERK;
	}

	public void setPRODUCT_MERK(String pRODUCT_MERK) {
		PRODUCT_MERK = pRODUCT_MERK;
	}

	public String getPRODUCT_TYPE() {
		return PRODUCT_TYPE;
	}

	public void setPRODUCT_TYPE(String pRODUCT_TYPE) {
		PRODUCT_TYPE = pRODUCT_TYPE;
	}

	public String getCATEGORY() {
		return CATEGORY;
	}

	public void setCATEGORY(String cATEGORY) {
		CATEGORY = cATEGORY;
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
