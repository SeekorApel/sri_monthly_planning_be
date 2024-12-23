package sri.sysint.sri_starter_back.model.view;

import java.math.BigDecimal;
import java.util.Date;

public class ViewDailyMonthlyPlanCuring {
    private BigDecimal detailDailyIdCuring;	
	private BigDecimal detailIdCuring;
    private Date dateDailyMp;
    private BigDecimal workDay;
	private BigDecimal totalPlan;
	private String description;
	private BigDecimal status;
	
	public ViewDailyMonthlyPlanCuring() {}

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

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public BigDecimal getStatus() {
		return status;
	}

	public void setStatus(BigDecimal status) {
		this.status = status;
	}

	public ViewDailyMonthlyPlanCuring(BigDecimal detailDailyIdCuring, BigDecimal detailIdCuring, Date dateDailyMp,
			BigDecimal workDay, BigDecimal totalPlan, String description, BigDecimal status) {
		super();
		this.detailDailyIdCuring = detailDailyIdCuring;
		this.detailIdCuring = detailIdCuring;
		this.dateDailyMp = dateDailyMp;
		this.workDay = workDay;
		this.totalPlan = totalPlan;
		this.description = description;
		this.status = status;
	}
	
	
}
