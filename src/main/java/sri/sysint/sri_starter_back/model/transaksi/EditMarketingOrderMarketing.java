package sri.sysint.sri_starter_back.model.transaksi;

import java.util.List;

import sri.sysint.sri_starter_back.model.DetailMarketingOrder;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.MarketingOrder;

public class EditMarketingOrderMarketing {
	private MarketingOrder marketingOrder;
    
    private List<HeaderMarketingOrder> headerMarketingOrder;
    
    private List<DetailMarketingOrder> detailMarketingOrder;
    
    public EditMarketingOrderMarketing() {
    	
    }
    
    public EditMarketingOrderMarketing(EditMarketingOrderMarketing editMarketingOrderMarketing) {
    	this.marketingOrder = editMarketingOrderMarketing.getMarketingOrder();
    	this.detailMarketingOrder = editMarketingOrderMarketing.getDetailMarketingOrder();
    	this.headerMarketingOrder = editMarketingOrderMarketing.getHeaderMarketingOrder();
    }

	public MarketingOrder getMarketingOrder() {
		return marketingOrder;
	}

	public void setMarketingOrder(MarketingOrder marketingOrder) {
		this.marketingOrder = marketingOrder;
	}

	public List<HeaderMarketingOrder> getHeaderMarketingOrder() {
		return headerMarketingOrder;
	}

	public void setHeaderMarketingOrder(List<HeaderMarketingOrder> headerMarketingOrder) {
		this.headerMarketingOrder = headerMarketingOrder;
	}

	public List<DetailMarketingOrder> getDetailMarketingOrder() {
		return detailMarketingOrder;
	}

	public void setDetailMarketingOrder(List<DetailMarketingOrder> detailMarketingOrder) {
		this.detailMarketingOrder = detailMarketingOrder;
	}
    
    
}
