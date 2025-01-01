package sri.sysint.sri_starter_back.model.view;

import java.util.Date;

public class ViewDistinctMarketingOrder {
	
    private Date month0;
    private Date month1;
    private Date month2;
    
    public ViewDistinctMarketingOrder() {
    	
    }
    
    public ViewDistinctMarketingOrder(Date month0, Date month1, Date month2) {
        this.month0 = month0;
        this.month1 = month1;
        this.month2 = month2;
    }

	public Date getMonth0() {
		return month0;
	}

	public void setMonth0(Date month0) {
		this.month0 = month0;
	}

	public Date getMonth1() {
		return month1;
	}

	public void setMonth1(Date month1) {
		this.month1 = month1;
	}

	public Date getMonth2() {
		return month2;
	}

	public void setMonth2(Date month2) {
		this.month2 = month2;
	}
    
}
