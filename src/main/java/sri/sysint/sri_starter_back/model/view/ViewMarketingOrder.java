package sri.sysint.sri_starter_back.model.view;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

public class ViewMarketingOrder {
	 	private String moId;
	    private String type;
	    private Date dateValid;
	    private BigDecimal status;
	    private BigDecimal revisionMarketing; 
	    private BigDecimal revisionPpc; 
	    private BigDecimal statusFilled;

	    
	    private List<ViewHeaderMarketingOrder> dataHeaderMo;
	    private List<ViewDetailMarketingOrder> dataDetailMo;
		public String getMoId() {
			return moId;
		}
		public void setMoId(String moId) {
			this.moId = moId;
		}
		public String getType() {
			return type;
		}
		public void setType(String type) {
			this.type = type;
		}
		public Date getDateValid() {
			return dateValid;
		}
		public void setDateValid(Date dateValid) {
			this.dateValid = dateValid;
		}
		
		public BigDecimal getStatusFilled() {
			return statusFilled;
		}
		public void setStatusFilled(BigDecimal statusFilled) {
			this.statusFilled = statusFilled;
		}
		public BigDecimal getStatus() {
			return status;
		}
		public void setStatus(BigDecimal status) {
			this.status = status;
		}
		public BigDecimal getRevisionMarketing() {
			return revisionMarketing;
		}
		public void setRevisionMarketing(BigDecimal revisionMarketing) {
			this.revisionMarketing = revisionMarketing;
		}
		public BigDecimal getRevisionPpc() {
			return revisionPpc;
		}
		public void setRevisionPpc(BigDecimal revisionPpc) {
			this.revisionPpc = revisionPpc;
		}
		public List<ViewHeaderMarketingOrder> getDataHeaderMo() {
			return dataHeaderMo;
		}
		public void setDataHeaderMo(List<ViewHeaderMarketingOrder> dataHeaderMo) {
			this.dataHeaderMo = dataHeaderMo;
		}
		public List<ViewDetailMarketingOrder> getDataDetailMo() {
			return dataDetailMo;
		}
		public void setDataDetailMo(List<ViewDetailMarketingOrder> dataDetailMo) {
			this.dataDetailMo = dataDetailMo;
		}
	    

		

}
