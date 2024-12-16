package sri.sysint.sri_starter_back.model.transaksi;

import java.util.List;

import sri.sysint.sri_starter_back.model.DetailMarketingOrder;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.MarketingOrder;

public class SaveFinalMarketingOrder {
	
	private MarketingOrder moFed;
	
	private List<HeaderMarketingOrder> headerMoFed;
	
	private List<DetailMarketingOrder> detailMoFed;
	
	private MarketingOrder moFdr;
	
	private List<HeaderMarketingOrder> headerMoFdr;
	
	private List<DetailMarketingOrder> detailMoFdr;
	
    public SaveFinalMarketingOrder() {
    	
    }
    
    public SaveFinalMarketingOrder(SaveFinalMarketingOrder mo) {
    	this.moFed = mo.getMoFed();
    	this.headerMoFed = mo.getHeaderMoFed();
    	this.detailMoFed = mo.getDetailMoFed();
    	this.moFdr = mo.getMoFdr();
    	this.headerMoFdr = mo.getHeaderMoFdr();
    	this.detailMoFdr = mo.getDetailMoFdr();
    }

	public MarketingOrder getMoFed() {
		return moFed;
	}

	public void setMoFed(MarketingOrder moFed) {
		this.moFed = moFed;
	}

	public List<HeaderMarketingOrder> getHeaderMoFed() {
		return headerMoFed;
	}

	public void setHeaderMoFed(List<HeaderMarketingOrder> headerMoFed) {
		this.headerMoFed = headerMoFed;
	}

	public List<DetailMarketingOrder> getDetailMoFed() {
		return detailMoFed;
	}

	public void setDetailMoFed(List<DetailMarketingOrder> detailMoFed) {
		this.detailMoFed = detailMoFed;
	}

	public MarketingOrder getMoFdr() {
		return moFdr;
	}

	public void setMoFdr(MarketingOrder moFdr) {
		this.moFdr = moFdr;
	}

	public List<HeaderMarketingOrder> getHeaderMoFdr() {
		return headerMoFdr;
	}

	public void setHeaderMoFdr(List<HeaderMarketingOrder> headerMoFdr) {
		this.headerMoFdr = headerMoFdr;
	}

	public List<DetailMarketingOrder> getDetailMoFdr() {
		return detailMoFdr;
	}

	public void setDetailMoFdr(List<DetailMarketingOrder> detailMoFdr) {
		this.detailMoFdr = detailMoFdr;
	}

}
