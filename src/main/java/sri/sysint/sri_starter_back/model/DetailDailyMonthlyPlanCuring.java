package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_D_DAILYMONTHLYPLAN_CURING")
public class DetailDailyMonthlyPlanCuring {
	
	@Id
    @Column(name = "DETAIL_DAILY_ID_CURING")
    private BigDecimal detailDailyIdCuring;
	
	@Column(name = "DETAIL_ID_CURING")
	private BigDecimal detailIdCuring;
	
    @Column(name = "DATE_DAILY_MP")
    private Date dateDailyMp;
    
    @Column(name = "WORK_DAY")
    private BigDecimal workDay;
	
	@Column(name = "TOTAL_PLAN")
	private BigDecimal totalPlan;
	
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
	
	public DetailDailyMonthlyPlanCuring() {
	}

	public DetailDailyMonthlyPlanCuring(BigDecimal detailDailyIdCuring, BigDecimal detailIdCuring, Date dateDailyMp,
			BigDecimal workDay, BigDecimal totalPlan, BigDecimal status, Date creationDate, String createdBy,
			Date lastUpdateDate, String lastUpdatedBy) {
		this.detailDailyIdCuring = detailDailyIdCuring;
		this.detailIdCuring = detailIdCuring;
		this.dateDailyMp = dateDailyMp;
		this.workDay = workDay;
		this.totalPlan = totalPlan;
		this.status = status;
		this.creationDate = creationDate;
		this.createdBy = createdBy;
		this.lastUpdateDate = lastUpdateDate;
		this.lastUpdatedBy = lastUpdatedBy;
	}
	
	public DetailDailyMonthlyPlanCuring(DetailDailyMonthlyPlanCuring detailDailyMonthlyPlanCuring) {
	    this.detailDailyIdCuring = detailDailyMonthlyPlanCuring.getDetailDailyIdCuring();
	    this.detailIdCuring = detailDailyMonthlyPlanCuring.getDetailIdCuring(); // Perbaiki nama properti dan gunakan getter
	    this.dateDailyMp = detailDailyMonthlyPlanCuring.getDateDailyMp(); // Gunakan getter untuk dateDailyMp
	    this.workDay = detailDailyMonthlyPlanCuring.getWorkDay(); // Gunakan getter untuk workDay
	    this.totalPlan = detailDailyMonthlyPlanCuring.getTotalPlan(); // Gunakan getter untuk totalPlan
	    this.status = detailDailyMonthlyPlanCuring.getStatus(); // Gunakan getter untuk status
	    this.creationDate = detailDailyMonthlyPlanCuring.getCreationDate(); // Gunakan getter untuk creationDate
	    this.createdBy = detailDailyMonthlyPlanCuring.getCreatedBy(); // Gunakan getter untuk createdBy
	    this.lastUpdateDate = detailDailyMonthlyPlanCuring.getLastUpdateDate(); // Gunakan getter untuk lastUpdateDate
	    this.lastUpdatedBy = detailDailyMonthlyPlanCuring.getLastUpdatedBy(); // Gunakan getter untuk lastUpdatedBy
	}

	public BigDecimal getDetailDailyIdCuring() {
		return detailDailyIdCuring;
	}

	public void setDetailDailyIdCuring(BigDecimal detailDailyIdCuring) {
		this.detailDailyIdCuring = detailDailyIdCuring;
	}

	
	public BigDecimal getDetailIdCuring() {
		return detailIdCuring;
	}

	public void setDetailIdCuring(BigDecimal detailIdCuring) {
		this.detailIdCuring = detailIdCuring;
	}

	public Date getDateDailyMp() {
		return dateDailyMp;
	}

	public void setDateDailyMp(Date dateDailyMp) {
		this.dateDailyMp = dateDailyMp;
	}

	public BigDecimal getWorkDay() {
		return workDay;
	}

	public void setWorkDay(BigDecimal workDay) {
		this.workDay = workDay;
	}

	public BigDecimal getTotalPlan() {
		return totalPlan;
	}

	public void setTotalPlan(BigDecimal totalPlan) {
		this.totalPlan = totalPlan;
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
	
	
}
