package sri.sysint.sri_starter_back.model.view;

import java.math.BigDecimal;

import javax.persistence.Column;

public class ViewDetailMarketingOrder {
	
	 private BigDecimal detailId;        // ubah dari idDetail
	    private String moId;
	    private String category;
	    private BigDecimal partNumber;
	    private String description;
	    private String machineType;
	    private BigDecimal capacity;        // ubah dari kapaPerMould
	    private BigDecimal qtyPerMould;     // ubah dari numberOfMould
	    private BigDecimal spareMould;
	    private BigDecimal mouldMonthlyPlan;
	    private BigDecimal qtyPerRak;
	    private BigDecimal minOrder;
	    private BigDecimal maxCapMonth0;    // ubah dari kapasitasMaksimum1
	    private BigDecimal maxCapMonth1;    // ubah dari kapasitasMaksimum2
	    private BigDecimal maxCapMonth2;    // ubah dari kapasitasMaksimum3

	    // input
	    private BigDecimal initialStock;    // ubah dari initialstock
	    private BigDecimal sfMonth0;        // ubah dari sf_month_0
	    private BigDecimal sfMonth1;        // ubah dari sf_month_1
	    private BigDecimal sfMonth2;        // ubah dari sf_month_2
	    private BigDecimal moMonth0;        // ubah dari mo_month_0
	    private BigDecimal moMonth1;        // ubah dari mo_month_1
	    private BigDecimal moMonth2;        // ubah dari mo_month_2

	    private BigDecimal ppd;
	    private BigDecimal cav;
	    private BigDecimal lockStatusM0;
	    private BigDecimal lockStatusM1;
	    private BigDecimal lockStatusM2;
		private BigDecimal ar;
		private BigDecimal defect;
		private BigDecimal reject;
	    private String itemCuring;

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
		public BigDecimal getSpareMould() {
			return spareMould;
		}
		public void setSpareMould(BigDecimal spareMould) {
			this.spareMould = spareMould;
		}
		public BigDecimal getMouldMonthlyPlan() {
			return mouldMonthlyPlan;
		}
		public void setMouldMonthlyPlan(BigDecimal mouldMonthlyPlan) {
			this.mouldMonthlyPlan = mouldMonthlyPlan;
		}
		public String getItemCuring() {
			return itemCuring;
		}
		public void setItemCuring(String itemCuring) {
			this.itemCuring = itemCuring;
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
		
		
		
}
