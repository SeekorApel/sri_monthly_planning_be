package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_M_PRODUCT")
public class Product {
	@Id
    @Column(name = "PART_NUMBER")
    private BigDecimal PART_NUMBER;
	
	@Column(name = "ITEM_CURING")
	private String ITEM_CURING;
	
	@Column(name = "PATTERN_ID")
	private BigDecimal PATTERN_ID;
	
	@Column(name = "SIZE_ID")
	private String SIZE_ID;
	
	@Column(name = "PRODUCT_TYPE_ID")
	private BigDecimal PRODUCT_TYPE_ID;
	
	@Column(name = "DESCRIPTION")
	private String DESCRIPTION;
	
	@Column(name = "RIM")
	private BigDecimal RIM;
	
	@Column(name = "WIB_TUBE")
	private String WIB_TUBE;
	
	@Column(name = "ITEM_ASSY")
	private String ITEM_ASSY;
	
	@Column(name = "ITEM_EXT")
	private String ITEM_EXT;
	
	@Column(name = "EXT_DESCRIPTION")
	private String EXT_DESCRIPTION;
	
	@Column(name = "QTY_PER_RAK")
	private BigDecimal QTY_PER_RAK;
	
	@Column(name = "UPPER_CONSTANT")
	private BigDecimal UPPER_CONSTANT;
	
	@Column(name = "LOWER_CONSTANT")
	private BigDecimal LOWER_CONSTANT;
	
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
	
	public Product() {}
	
	public Product(Product product) {
		this.PART_NUMBER = product.getPART_NUMBER();
		this.ITEM_CURING = product.getITEM_CURING();
		this.PATTERN_ID = product.getPATTERN_ID();
		this.SIZE_ID = product.getSIZE_ID();
		this.PRODUCT_TYPE_ID = product.getPRODUCT_TYPE_ID();
		this.DESCRIPTION = product.getDESCRIPTION();
		this.RIM = product.getRIM();
		this.WIB_TUBE = product.getWIB_TUBE();
		this.ITEM_ASSY = product.getITEM_ASSY();
		this.ITEM_EXT = product.getITEM_EXT();
		this.EXT_DESCRIPTION = product.getEXT_DESCRIPTION();
		this.QTY_PER_RAK = product.getQTY_PER_RAK();
		this.UPPER_CONSTANT = product.getUPPER_CONSTANT();
		this.LOWER_CONSTANT = product.getLOWER_CONSTANT();
		this.STATUS = product.getSTATUS();
		this.CREATION_DATE = product.getCREATION_DATE();
		this.CREATED_BY = product.getCREATED_BY();
		this.LAST_UPDATE_DATE = product.getLAST_UPDATE_DATE();
		this.LAST_UPDATED_BY = product.getLAST_UPDATED_BY();
	}
	
	

	public Product(BigDecimal pART_NUMBER, String iTEM_CURING, BigDecimal pATTERN_ID, String sIZE_ID,
			BigDecimal pRODUCT_TYPE_ID, String dESCRIPTION, BigDecimal rIM, String wIB_TUBE, String iTEM_ASSY,
			String iTEM_EXT, String eXT_DESCRIPTION, BigDecimal qTY_PER_RAK, BigDecimal uPPER_CONSTANT,
			BigDecimal lOWER_CONSTANT, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY, Date lAST_UPDATE_DATE,
			String lAST_UPDATED_BY) {
		super();
		PART_NUMBER = pART_NUMBER;
		ITEM_CURING = iTEM_CURING;
		PATTERN_ID = pATTERN_ID;
		SIZE_ID = sIZE_ID;
		PRODUCT_TYPE_ID = pRODUCT_TYPE_ID;
		DESCRIPTION = dESCRIPTION;
		RIM = rIM;
		WIB_TUBE = wIB_TUBE;
		ITEM_ASSY = iTEM_ASSY;
		ITEM_EXT = iTEM_EXT;
		EXT_DESCRIPTION = eXT_DESCRIPTION;
		QTY_PER_RAK = qTY_PER_RAK;
		UPPER_CONSTANT = uPPER_CONSTANT;
		LOWER_CONSTANT = lOWER_CONSTANT;
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

	public BigDecimal getPATTERN_ID() {
		return PATTERN_ID;
	}

	public void setPATTERN_ID(BigDecimal pATTERN_ID) {
		PATTERN_ID = pATTERN_ID;
	}

	public String getSIZE_ID() {
		return SIZE_ID;
	}

	public void setSIZE_ID(String sIZE_ID) {
		SIZE_ID = sIZE_ID;
	}

	public BigDecimal getPRODUCT_TYPE_ID() {
		return PRODUCT_TYPE_ID;
	}

	public void setPRODUCT_TYPE_ID(BigDecimal pRODUCT_TYPE_ID) {
		PRODUCT_TYPE_ID = pRODUCT_TYPE_ID;
	}

	public String getDESCRIPTION() {
		return DESCRIPTION;
	}

	public void setDESCRIPTION(String dESCRIPTION) {
		DESCRIPTION = dESCRIPTION;
	}

	public BigDecimal getRIM() {
		return RIM;
	}

	public void setRIM(BigDecimal rIM) {
		RIM = rIM;
	}

	public String getWIB_TUBE() {
		return WIB_TUBE;
	}

	public void setWIB_TUBE(String wIB_TUBE) {
		WIB_TUBE = wIB_TUBE;
	}

	public String getITEM_ASSY() {
		return ITEM_ASSY;
	}

	public void setITEM_ASSY(String iTEM_ASSY) {
		ITEM_ASSY = iTEM_ASSY;
	}

	public String getITEM_EXT() {
		return ITEM_EXT;
	}

	public void setITEM_EXT(String iTEM_EXT) {
		ITEM_EXT = iTEM_EXT;
	}

	public String getEXT_DESCRIPTION() {
		return EXT_DESCRIPTION;
	}

	public void setEXT_DESCRIPTION(String eXT_DESCRIPTION) {
		EXT_DESCRIPTION = eXT_DESCRIPTION;
	}

	public BigDecimal getQTY_PER_RAK() {
		return QTY_PER_RAK;
	}

	public void setQTY_PER_RAK(BigDecimal qTY_PER_RAK) {
		QTY_PER_RAK = qTY_PER_RAK;
	}

	public BigDecimal getUPPER_CONSTANT() {
		return UPPER_CONSTANT;
	}

	public void setUPPER_CONSTANT(BigDecimal uPPER_CONSTANT) {
		UPPER_CONSTANT = uPPER_CONSTANT;
	}

	public BigDecimal getLOWER_CONSTANT() {
		return LOWER_CONSTANT;
	}

	public void setLOWER_CONSTANT(BigDecimal lOWER_CONSTANT) {
		LOWER_CONSTANT = lOWER_CONSTANT;
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
