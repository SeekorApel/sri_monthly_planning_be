package sri.sysint.sri_starter_back.model.transaksi;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import sri.sysint.sri_starter_back.model.MonthlyPlanningCuring;
import sri.sysint.sri_starter_back.model.DetailMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.DetailDailyMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.ShiftMonthlyPlan;
import sri.sysint.sri_starter_back.service.MonthlyPlanServiceImpl.ChangeMould;


public class ViewMonthlyPlanning {
	
    private MonthlyPlanningCuring monthlyPlanningCuring;
    
    private List<DetailMonthlyPlanCuring> detailMonthlyPlanCuring;
    
    private List<DetailDailyMonthlyPlanCuring> detailDailyMonthlyPlanCuring;
    
    private List<ShiftMonthlyPlan> shiftMonthlyPlan;
    
    private List<ChangeMould> changeMould;
    
    List<Map<String, Object>> productDetails = new ArrayList<>();

    
    private List<Map<String, Object>> sizePa;
    private List<Map<String, Object>> description;

    public ViewMonthlyPlanning() {}

	public ViewMonthlyPlanning(MonthlyPlanningCuring monthlyPlanningCuring,
			List<DetailMonthlyPlanCuring> detailMonthlyPlanCuring,
			List<DetailDailyMonthlyPlanCuring> detailDailyMonthlyPlanCuring, List<ShiftMonthlyPlan> shiftMonthlyPlan,
			List<ChangeMould> changeMould, List<Map<String, Object>> productDetails, List<Map<String, Object>> sizePa,
			List<Map<String, Object>> description) {
		super();
		this.monthlyPlanningCuring = monthlyPlanningCuring;
		this.detailMonthlyPlanCuring = detailMonthlyPlanCuring;
		this.detailDailyMonthlyPlanCuring = detailDailyMonthlyPlanCuring;
		this.shiftMonthlyPlan = shiftMonthlyPlan;
		this.changeMould = changeMould;
		this.productDetails = productDetails;
		this.sizePa = sizePa;
		this.description = description;
	}

	public MonthlyPlanningCuring getMonthlyPlanningCuring() {
		return monthlyPlanningCuring;
	}

	public void setMonthlyPlanningCuring(MonthlyPlanningCuring monthlyPlanningCuring) {
		this.monthlyPlanningCuring = monthlyPlanningCuring;
	}

	public List<DetailMonthlyPlanCuring> getDetailMonthlyPlanCuring() {
		return detailMonthlyPlanCuring;
	}

	public void setDetailMonthlyPlanCuring(List<DetailMonthlyPlanCuring> detailMonthlyPlanCuring) {
		this.detailMonthlyPlanCuring = detailMonthlyPlanCuring;
	}

	public List<DetailDailyMonthlyPlanCuring> getDetailDailyMonthlyPlanCuring() {
		return detailDailyMonthlyPlanCuring;
	}

	public void setDetailDailyMonthlyPlanCuring(List<DetailDailyMonthlyPlanCuring> detailDailyMonthlyPlanCuring) {
		this.detailDailyMonthlyPlanCuring = detailDailyMonthlyPlanCuring;
	}

	public List<ShiftMonthlyPlan> getShiftMonthlyPlan() {
		return shiftMonthlyPlan;
	}

	public void setShiftMonthlyPlan(List<ShiftMonthlyPlan> shiftMonthlyPlan) {
		this.shiftMonthlyPlan = shiftMonthlyPlan;
	}

	public List<ChangeMould> getChangeMould() {
		return changeMould;
	}

	public void setChangeMould(List<ChangeMould> changeMould) {
		this.changeMould = changeMould;
	}

	public List<Map<String, Object>> getProductDetails() {
		return productDetails;
	}

	public void setProductDetails(List<Map<String, Object>> productDetails) {
		this.productDetails = productDetails;
	}

	public List<Map<String, Object>> getSizePa() {
		return sizePa;
	}

	public void setSizePa(List<Map<String, Object>> sizePa) {
		this.sizePa = sizePa;
	}

	public List<Map<String, Object>> getDescription() {
		return description;
	}

	public void setDescription(List<Map<String, Object>> description) {
		this.description = description;
	}

	

}
