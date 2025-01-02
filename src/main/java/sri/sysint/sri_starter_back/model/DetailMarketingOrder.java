package sri.sysint.sri_starter_back.model;

import javax.persistence.Entity;
import javax.persistence.Table;
import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Id;

@Entity
@Table(name = "SRI_IMPP_D_MARKETINGORDER")
public class DetailMarketingOrder {
	
	@Id
    @Column(name = "DETAIL_ID")
    private BigDecimal detailId;
	
	@Column(name = "MO_ID")
	private String moId;
	
	@Column(name = "CATEGORY")
	private String category;
	
	@Column(name = "PART_NUMBER")
	private BigDecimal partNumber;
	
	@Column(name = "DESCRIPTION")
	private String description;
	
	@Column(name = "MACHINE_TYPE")
	private String machineType;
	
	@Column(name = "CAPACITY")
	private BigDecimal capacity;
	
	@Column(name = "QTY_PER_MOULD")
	private BigDecimal qtyPerMould;
	
	@Column(name = "QTY_PER_RAK")
	private BigDecimal qtyPerRak;
	
	@Column(name = "MIN_ORDER")
	private BigDecimal minOrder;
	
	@Column(name = "MAX_CAP_MONTH_0")
	private BigDecimal maxCapMonth0;
	
	@Column(name = "MAX_CAP_MONTH_1")
	private BigDecimal maxCapMonth1;
	
	@Column(name = "MAX_CAP_MONTH_2")
	private BigDecimal maxCapMonth2;
	
	@Column(name = "INITIAL_STOCK")
	private BigDecimal initialStock;
	
	@Column(name = "SF_MONTH_0")
	private BigDecimal sfMonth0;
	
	@Column(name = "SF_MONTH_1")
	private BigDecimal sfMonth1;
	
	@Column(name = "SF_MONTH_2")
	private BigDecimal sfMonth2;
	
	@Column(name = "MO_MONTH_0")
	private BigDecimal moMonth0;
	
	@Column(name = "MO_MONTH_1")
	private BigDecimal moMonth1;
	
	@Column(name = "MO_MONTH_2")
	private BigDecimal moMonth2;
	
	@Column(name = "PPD")
	private BigDecimal ppd;
	
	@Column(name = "CAV")
	private BigDecimal cav;
	
	@Column(name = "LOCK_STATUS_MONTH_0")
	private BigDecimal lockStatusM0;
	
	@Column(name = "LOCK_STATUS_MONTH_1")
	private BigDecimal lockStatusM1;
	
	@Column(name = "LOCK_STATUS_MONTH_2")
	private BigDecimal lockStatusM2;
		
	@Column(name = "AR")
	private BigDecimal ar;
	
	@Column(name = "DEFECT")
	private BigDecimal defect;
	
	@Column(name = "REJECT")
	private BigDecimal reject;
	
	@Column(name = "TOTAL_AR")
	private BigDecimal totalAr;
	
	@Column(name = "STATUS")
	private BigDecimal status;
	
	@Column(name = "CREATION_DATE")
	private Date creationDate;
	
	@Column(name = "CREATED_BY")
	private String createdBy;
	
	@Column(name = "LAST_UPDATE_DATE")
	private Date lastUpdateDate;
	
	@Column(name = "LAST_UPDATED_BY")
	private String lastUpdatedBy;

	public DetailMarketingOrder(DetailMarketingOrder detailMarketingOrder) {
	    this.detailId = detailMarketingOrder.getDetailId();
	    this.moId = detailMarketingOrder.getMoId();
	    this.category = detailMarketingOrder.getCategory();
	    this.partNumber = detailMarketingOrder.getPartNumber();
	    this.description = detailMarketingOrder.getDescription();
	    this.machineType = detailMarketingOrder.getMachineType();
	    this.capacity = detailMarketingOrder.getCapacity();
	    this.qtyPerMould = detailMarketingOrder.getQtyPerMould();
	    this.qtyPerRak = detailMarketingOrder.getQtyPerRak();
	    this.minOrder = detailMarketingOrder.getMinOrder();
	    this.maxCapMonth0 = detailMarketingOrder.getMaxCapMonth0();
	    this.maxCapMonth1 = detailMarketingOrder.getMaxCapMonth1();
	    this.maxCapMonth2 = detailMarketingOrder.getMaxCapMonth2();
	    this.initialStock = detailMarketingOrder.getInitialStock();
	    this.sfMonth0 = detailMarketingOrder.getSfMonth0();
	    this.sfMonth1 = detailMarketingOrder.getSfMonth1();
	    this.sfMonth2 = detailMarketingOrder.getSfMonth2();
	    this.moMonth0 = detailMarketingOrder.getMoMonth0();
	    this.moMonth1 = detailMarketingOrder.getMoMonth1();
	    this.moMonth2 = detailMarketingOrder.getMoMonth2();
	    this.ppd = detailMarketingOrder.getPpd();
	    this.cav = detailMarketingOrder.getCav();
	    this.lockStatusM0 = detailMarketingOrder.getLockStatusM0();
	    this.lockStatusM1 = detailMarketingOrder.getLockStatusM1();
	    this.lockStatusM2 = detailMarketingOrder.getLockStatusM2();
	    this.totalAr = detailMarketingOrder.getTotalAr();
	    this.ar = detailMarketingOrder.getAr();
	    this.defect = detailMarketingOrder.getDefect();
	    this.reject = detailMarketingOrder.getReject();
	    this.totalAr = detailMarketingOrder.getTotalAr();
	    this.status = detailMarketingOrder.getStatus();
	    this.creationDate = detailMarketingOrder.getCreationDate();
	    this.createdBy = detailMarketingOrder.getCreatedBy();
	    this.lastUpdateDate = detailMarketingOrder.getLastUpdateDate();
	    this.lastUpdatedBy = detailMarketingOrder.getLastUpdatedBy();
	}

	public DetailMarketingOrder() {}

	
	public BigDecimal getDetailId() {
		return detailId;
	}

	public void setDetailId(BigDecimal detailId) {
		this.detailId = detailId;
	}

	public String getMoId() {
		return moId;
	}

	public void setMoId(String moId) {
		this.moId = moId;
	}

	public String getCategory() {
		return category;
	}

	public void setCategory(String category) {
		this.category = category;
	}


	public BigDecimal getPartNumber() {
		return partNumber;
	}

	public void setPartNumber(BigDecimal partNumber) {
		this.partNumber = partNumber;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getMachineType() {
		return machineType;
	}

	public void setMachineType(String machineType) {
		this.machineType = machineType;
	}

	public BigDecimal getCapacity() {
		return capacity;
	}

	public void setCapacity(BigDecimal capacity) {
		this.capacity = capacity;
	}

	public BigDecimal getQtyPerMould() {
		return qtyPerMould;
	}

	public void setQtyPerMould(BigDecimal qtyPerMould) {
		this.qtyPerMould = qtyPerMould;
	}

	public BigDecimal getQtyPerRak() {
		return qtyPerRak;
	}

	public void setQtyPerRak(BigDecimal qtyPerRak) {
		this.qtyPerRak = qtyPerRak;
	}

	public BigDecimal getMinOrder() {
		return minOrder;
	}

	public void setMinOrder(BigDecimal minOrder) {
		this.minOrder = minOrder;
	}

	public BigDecimal getMaxCapMonth0() {
		return maxCapMonth0;
	}

	public void setMaxCapMonth0(BigDecimal maxCapMonth0) {
		this.maxCapMonth0 = maxCapMonth0;
	}

	public BigDecimal getMaxCapMonth1() {
		return maxCapMonth1;
	}

	public void setMaxCapMonth1(BigDecimal maxCapMonth1) {
		this.maxCapMonth1 = maxCapMonth1;
	}

	public BigDecimal getMaxCapMonth2() {
		return maxCapMonth2;
	}

	public void setMaxCapMonth2(BigDecimal maxCapMonth2) {
		this.maxCapMonth2 = maxCapMonth2;
	}

	public BigDecimal getInitialStock() {
		return initialStock;
	}

	public void setInitialStock(BigDecimal initialStock) {
		this.initialStock = initialStock;
	}

	public BigDecimal getSfMonth0() {
		return sfMonth0;
	}

	public void setSfMonth0(BigDecimal sfMonth0) {
		this.sfMonth0 = sfMonth0;
	}

	public BigDecimal getSfMonth1() {
		return sfMonth1;
	}

	public void setSfMonth1(BigDecimal sfMonth1) {
		this.sfMonth1 = sfMonth1;
	}

	public BigDecimal getSfMonth2() {
		return sfMonth2;
	}

	public void setSfMonth2(BigDecimal sfMonth2) {
		this.sfMonth2 = sfMonth2;
	}

	public BigDecimal getMoMonth0() {
		return moMonth0;
	}

	public void setMoMonth0(BigDecimal moMonth0) {
		this.moMonth0 = moMonth0;
	}

	public BigDecimal getMoMonth1() {
		return moMonth1;
	}

	public void setMoMonth1(BigDecimal moMonth1) {
		this.moMonth1 = moMonth1;
	}

	public BigDecimal getMoMonth2() {
		return moMonth2;
	}

	public void setMoMonth2(BigDecimal moMonth2) {
		this.moMonth2 = moMonth2;
	}

	public BigDecimal getPpd() {
		return ppd;
	}

	public void setPpd(BigDecimal ppd) {
		this.ppd = ppd;
	}

	public BigDecimal getCav() {
		return cav;
	}

	public void setCav(BigDecimal cav) {
		this.cav = cav;
	}


	public BigDecimal getLockStatusM0() {
		return lockStatusM0;
	}

	public void setLockStatusM0(BigDecimal lockStatusM0) {
		this.lockStatusM0 = lockStatusM0;
	}

	public BigDecimal getLockStatusM1() {
		return lockStatusM1;
	}

	public void setLockStatusM1(BigDecimal lockStatusM1) {
		this.lockStatusM1 = lockStatusM1;
	}

	public BigDecimal getLockStatusM2() {
		return lockStatusM2;
	}

	public void setLockStatusM2(BigDecimal lockStatusM2) {
		this.lockStatusM2 = lockStatusM2;
	}
	
	
	public BigDecimal getTotalAr() {
		return totalAr;
	}

	public void setTotalAr(BigDecimal totalAr) {
		this.totalAr = totalAr;
	}

	public BigDecimal getAr() {
		return ar;
	}

	public void setAr(BigDecimal ar) {
		this.ar = ar;
	}

	public BigDecimal getDefect() {
		return defect;
	}

	public void setDefect(BigDecimal defect) {
		this.defect = defect;
	}

	public BigDecimal getReject() {
		return reject;
	}

	public void setReject(BigDecimal reject) {
		this.reject = reject;
	}

	public BigDecimal getStatus() {
		return status;
	}

	public void setStatus(BigDecimal status) {
		this.status = status;
	}

	public Date getCreationDate() {
		return creationDate;
	}

	public void setCreationDate(Date creationDate) {
		this.creationDate = creationDate;
	}

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}

	public Date getLastUpdateDate() {
		return lastUpdateDate;
	}

	public void setLastUpdateDate(Date lastUpdateDate) {
		this.lastUpdateDate = lastUpdateDate;
	}

	public String getLastUpdatedBy() {
		return lastUpdatedBy;
	}

	public void setLastUpdatedBy(String lastUpdatedBy) {
		this.lastUpdatedBy = lastUpdatedBy;
	}

	@Override
	public String toString() {
		return "DetailMarketingOrder [moId=" + moId + ", partNumber=" + partNumber + ", description=" + description
				+ ", moMonth0=" + moMonth0 + ", ppd=" + ppd + ", cav=" + cav + ", lockStatusM0=" + lockStatusM0
				+ ", lockStatusM1=" + lockStatusM1 + ", lockStatusM2=" + lockStatusM2 + ", ar=" + ar + ", defect="
				+ defect + ", reject=" + reject + "]";
	}
	
	

	
}
