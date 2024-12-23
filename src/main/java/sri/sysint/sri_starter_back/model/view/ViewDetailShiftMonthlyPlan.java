package sri.sysint.sri_starter_back.model.view;

import java.math.BigDecimal;
import java.util.Date;

public class ViewDetailShiftMonthlyPlan {
    private BigDecimal detailShiftIdCuring;
    private BigDecimal detailDailyIdCuring;
    private String moId;
    private String mpCuringId;
    private Date actualDate;
    private String itemCuring;
    private BigDecimal partNumber;
    private String workCenterText;
    private BigDecimal cavity;
    private BigDecimal cavityUsage;
    private BigDecimal whShift1;
    private BigDecimal whShift2;
    private BigDecimal whShift3;
    private BigDecimal kapaShift1;
    private BigDecimal kapaShift2;
    private BigDecimal kapaShift3;
    private BigDecimal totalKapa;
    private BigDecimal cavityExist;
    private Date changeDate;
    private String wct;
    private BigDecimal shift;

    
    public ViewDetailShiftMonthlyPlan() {}

	public ViewDetailShiftMonthlyPlan(BigDecimal detailShiftIdCuring, BigDecimal detailDailyIdCuring, String moId,
			String mpCuringId, Date actualDate, String itemCuring, BigDecimal partNumber, String workCenterText,
			BigDecimal cavity, BigDecimal cavityUsage, BigDecimal whShift1, BigDecimal whShift2, BigDecimal whShift3,
			BigDecimal kapaShift1, BigDecimal kapaShift2, BigDecimal kapaShift3, BigDecimal totalKapa,
			BigDecimal cavityExist, Date changeDate, String wct, BigDecimal shift) {
		super();
		this.detailShiftIdCuring = detailShiftIdCuring;
		this.detailDailyIdCuring = detailDailyIdCuring;
		this.moId = moId;
		this.mpCuringId = mpCuringId;
		this.actualDate = actualDate;
		this.itemCuring = itemCuring;
		this.partNumber = partNumber;
		this.workCenterText = workCenterText;
		this.cavity = cavity;
		this.cavityUsage = cavityUsage;
		this.whShift1 = whShift1;
		this.whShift2 = whShift2;
		this.whShift3 = whShift3;
		this.kapaShift1 = kapaShift1;
		this.kapaShift2 = kapaShift2;
		this.kapaShift3 = kapaShift3;
		this.totalKapa = totalKapa;
		this.cavityExist = cavityExist;
		this.changeDate = changeDate;
		this.wct = wct;
		this.shift = shift;
	}


	public BigDecimal getDetailShiftIdCuring() {
		return detailShiftIdCuring;
	}


	public void setDetailShiftIdCuring(BigDecimal detailShiftIdCuring) {
		this.detailShiftIdCuring = detailShiftIdCuring;
	}


	public BigDecimal getDetailDailyIdCuring() {
		return detailDailyIdCuring;
	}


	public void setDetailDailyIdCuring(BigDecimal detailDailyIdCuring) {
		this.detailDailyIdCuring = detailDailyIdCuring;
	}


	public String getMoId() {
		return moId;
	}


	public void setMoId(String moId) {
		this.moId = moId;
	}


	public String getMpCuringId() {
		return mpCuringId;
	}


	public void setMpCuringId(String mpCuringId) {
		this.mpCuringId = mpCuringId;
	}


	public Date getActualDate() {
		return actualDate;
	}


	public void setActualDate(Date actualDate) {
		this.actualDate = actualDate;
	}


	public String getItemCuring() {
		return itemCuring;
	}


	public void setItemCuring(String itemCuring) {
		this.itemCuring = itemCuring;
	}


	public BigDecimal getPartNumber() {
		return partNumber;
	}


	public void setPartNumber(BigDecimal partNumber) {
		this.partNumber = partNumber;
	}


	public String getWorkCenterText() {
		return workCenterText;
	}


	public void setWorkCenterText(String workCenterText) {
		this.workCenterText = workCenterText;
	}


	public BigDecimal getCavity() {
		return cavity;
	}


	public void setCavity(BigDecimal cavity) {
		this.cavity = cavity;
	}


	public BigDecimal getCavityUsage() {
		return cavityUsage;
	}


	public void setCavityUsage(BigDecimal cavityUsage) {
		this.cavityUsage = cavityUsage;
	}


	public BigDecimal getWhShift1() {
		return whShift1;
	}


	public void setWhShift1(BigDecimal whShift1) {
		this.whShift1 = whShift1;
	}


	public BigDecimal getWhShift2() {
		return whShift2;
	}


	public void setWhShift2(BigDecimal whShift2) {
		this.whShift2 = whShift2;
	}


	public BigDecimal getWhShift3() {
		return whShift3;
	}


	public void setWhShift3(BigDecimal whShift3) {
		this.whShift3 = whShift3;
	}


	public BigDecimal getKapaShift1() {
		return kapaShift1;
	}


	public void setKapaShift1(BigDecimal kapaShift1) {
		this.kapaShift1 = kapaShift1;
	}


	public BigDecimal getKapaShift2() {
		return kapaShift2;
	}


	public void setKapaShift2(BigDecimal kapaShift2) {
		this.kapaShift2 = kapaShift2;
	}


	public BigDecimal getKapaShift3() {
		return kapaShift3;
	}


	public void setKapaShift3(BigDecimal kapaShift3) {
		this.kapaShift3 = kapaShift3;
	}


	public BigDecimal getTotalKapa() {
		return totalKapa;
	}


	public void setTotalKapa(BigDecimal totalKapa) {
		this.totalKapa = totalKapa;
	}


	public BigDecimal getCavityExist() {
		return cavityExist;
	}


	public void setCavityExist(BigDecimal cavityExist) {
		this.cavityExist = cavityExist;
	}


	public Date getChangeDate() {
		return changeDate;
	}


	public void setChangeDate(Date changeDate) {
		this.changeDate = changeDate;
	}


	public String getWct() {
		return wct;
	}


	public void setWct(String wct) {
		this.wct = wct;
	}


	public BigDecimal getShift() {
		return shift;
	}


	public void setShift(BigDecimal shift) {
		this.shift = shift;
	}
    
}
