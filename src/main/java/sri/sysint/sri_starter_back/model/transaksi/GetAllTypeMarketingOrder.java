package sri.sysint.sri_starter_back.model.transaksi;
import java.util.List;
import sri.sysint.sri_starter_back.model.DetailMarketingOrder;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.MarketingOrder;

public class GetAllTypeMarketingOrder {
	
	private MarketingOrder moFed;
    
    private List<HeaderMarketingOrder> headerMarketingOrderFed;
    
    private List<DetailMarketingOrder> detailMarketingOrderFed;
    
	private MarketingOrder moFdr;
    
    private List<HeaderMarketingOrder> headerMarketingOrderFdr;
    
    private List<DetailMarketingOrder> detailMarketingOrderFdr;
    
    public GetAllTypeMarketingOrder() {
    	
    }

	public GetAllTypeMarketingOrder(GetAllTypeMarketingOrder getAllTypeMarketingOrder) {
		this.moFed = getAllTypeMarketingOrder.getMoFed();
		this.headerMarketingOrderFed = getAllTypeMarketingOrder.getHeaderMarketingOrderFed();
		this.detailMarketingOrderFed = getAllTypeMarketingOrder.getDetailMarketingOrderFed();
		this.moFdr = getAllTypeMarketingOrder.getMoFdr();
		this.headerMarketingOrderFdr = getAllTypeMarketingOrder.getHeaderMarketingOrderFdr();
		this.detailMarketingOrderFdr = getAllTypeMarketingOrder.getDetailMarketingOrderFdr();
	}

	public MarketingOrder getMoFed() {
		return moFed;
	}

	public void setMoFed(MarketingOrder moFed) {
		this.moFed = moFed;
	}

	public List<HeaderMarketingOrder> getHeaderMarketingOrderFed() {
		return headerMarketingOrderFed;
	}

	public void setHeaderMarketingOrderFed(List<HeaderMarketingOrder> headerMarketingOrderFed) {
		this.headerMarketingOrderFed = headerMarketingOrderFed;
	}

	public List<DetailMarketingOrder> getDetailMarketingOrderFed() {
		return detailMarketingOrderFed;
	}

	public void setDetailMarketingOrderFed(List<DetailMarketingOrder> detailMarketingOrderFed) {
		this.detailMarketingOrderFed = detailMarketingOrderFed;
	}

	public MarketingOrder getMoFdr() {
		return moFdr;
	}

	public void setMoFdr(MarketingOrder moFdr) {
		this.moFdr = moFdr;
	}

	public List<HeaderMarketingOrder> getHeaderMarketingOrderFdr() {
		return headerMarketingOrderFdr;
	}

	public void setHeaderMarketingOrderFdr(List<HeaderMarketingOrder> headerMarketingOrderFdr) {
		this.headerMarketingOrderFdr = headerMarketingOrderFdr;
	}

	public List<DetailMarketingOrder> getDetailMarketingOrderFdr() {
		return detailMarketingOrderFdr;
	}

	public void setDetailMarketingOrderFdr(List<DetailMarketingOrder> detailMarketingOrderFdr) {
		this.detailMarketingOrderFdr = detailMarketingOrderFdr;
	}
    
    
    

}
