package sri.sysint.sri_starter_back.model;

import javax.persistence.Entity;
import javax.persistence.Table;
import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Id;

@Entity
@Table(name = "SRI_IMPP_D_HEADERMARKETINGORDER")
public class HeaderMarketingOrder {

	@Id
	@Column(name = "HEADER_ID")
	private BigDecimal headerId;

	@Column(name = "MO_ID")
	private String moId;

	@Column(name = "MONTH")
	private Date month;

	@Column(name = "WD_NORMAL_TIRE")
	private BigDecimal wdNormalTire;
	
	@Column(name = "WD_OT_TL")
	private BigDecimal wdOtTl;
	
	@Column(name = "WD_OT_TT")
	private BigDecimal wdOtTt;
	
	@Column(name = "WD_NORMAL_TUBE")
	private BigDecimal wdNormalTube;
	
	@Column(name = "WD_OT_TUBE")
	private BigDecimal wdOtTube;
	
	@Column(name = "TOTAL_WD_TL")
	private BigDecimal totalWdTl;
	
	@Column(name = "TOTAL_WD_TT")
	private BigDecimal totalWdTt;
	
	@Column(name = "TOTAL_WD_TUBE")
	private BigDecimal totalWdTube;
	
	@Column(name = "MAX_CAP_TUBE")
	private BigDecimal maxCapTube;
	
	@Column(name = "MAX_CAP_TL")
	private BigDecimal maxCapTl;
	
	@Column(name = "MAX_CAP_TT")
	private BigDecimal maxCapTt;
	
	@Column(name = "AIRBAG_MACHINE")
	private BigDecimal airbagMachine;
	
	@Column(name = "TL")
	private BigDecimal tl;
	
	@Column(name = "TT")
	private BigDecimal tt;
	
	@Column(name = "TOTAL_MO")
	private BigDecimal totalMo;
	
	@Column(name = "TL_PERCENTAGE")
	private BigDecimal tlPercentage;
	
	@Column(name = "TT_PERCENTAGE")
	private BigDecimal ttPercentage;
	
	@Column(name = "NOTE_ORDER_TL")
	private String noteOrderTl;
	
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
	
	public HeaderMarketingOrder() {}
	
	 public HeaderMarketingOrder(HeaderMarketingOrder headerMarketingOrder) {
	        this.headerId = headerMarketingOrder.getHeaderId();
	        this.moId = headerMarketingOrder.getMoId();
	        this.month = headerMarketingOrder.getMonth();
	        this.wdNormalTire = headerMarketingOrder.getWdNormalTire();
	        this.wdOtTl = headerMarketingOrder.getWdOtTl();
	        this.wdOtTt = headerMarketingOrder.getWdOtTt();
	        this.wdNormalTube = headerMarketingOrder.getWdNormalTube();
	        this.wdOtTube = headerMarketingOrder.getWdOtTube();
	        this.totalWdTl = headerMarketingOrder.getTotalWdTl();
	        this.totalWdTt = headerMarketingOrder.getTotalWdTt();
	        this.totalWdTube = headerMarketingOrder.getTotalWdTube();
	        this.maxCapTube = headerMarketingOrder.getMaxCapTube();
	        this.maxCapTl = headerMarketingOrder.getMaxCapTl();
	        this.maxCapTt = headerMarketingOrder.getMaxCapTt();
	        this.airbagMachine = headerMarketingOrder.getAirbagMachine();
	        this.tl = headerMarketingOrder.getTl();
	        this.tt = headerMarketingOrder.getTt();
	        this.totalMo = headerMarketingOrder.getTotalMo();
	        this.tlPercentage = headerMarketingOrder.getTlPercentage();
	        this.ttPercentage = headerMarketingOrder.getTtPercentage();
	        this.noteOrderTl = headerMarketingOrder.getNoteOrderTl();
	        this.status = headerMarketingOrder.getStatus();
	        this.creationDate = headerMarketingOrder.getCreationDate();
	        this.createdBy = headerMarketingOrder.getCreatedBy();
	        this.lastUpdateDate = headerMarketingOrder.getLastUpdateDate();
	        this.lastUpdatedBy = headerMarketingOrder.getLastUpdatedBy();
	    }


	public HeaderMarketingOrder(BigDecimal headerId, String moId, Date month, BigDecimal wdNormalTire,
			BigDecimal wdOtTl, BigDecimal wdOtTt, BigDecimal wdNormalTube, BigDecimal wdOtTube, BigDecimal totalWdTl,
			BigDecimal totalWdTt, BigDecimal totalWdTube, BigDecimal maxCapTube, BigDecimal maxCapTl,
			BigDecimal maxCapTt, BigDecimal airbagMachine, BigDecimal tl, BigDecimal tt, BigDecimal totalMo,
			BigDecimal tlPercentage, BigDecimal ttPercentage, String noteOrderTl, BigDecimal status, Date creationDate,
			String createdBy, Date lastUpdateDate, String lastUpdatedBy) {
		super();
		this.headerId = headerId;
		this.moId = moId;
		this.month = month;
		this.wdNormalTire = wdNormalTire;
		this.wdOtTl = wdOtTl;
		this.wdOtTt = wdOtTt;
		this.wdNormalTube = wdNormalTube;
		this.wdOtTube = wdOtTube;
		this.totalWdTl = totalWdTl;
		this.totalWdTt = totalWdTt;
		this.totalWdTube = totalWdTube;
		this.maxCapTube = maxCapTube;
		this.maxCapTl = maxCapTl;
		this.maxCapTt = maxCapTt;
		this.airbagMachine = airbagMachine;
		this.tl = tl;
		this.tt = tt;
		this.totalMo = totalMo;
		this.tlPercentage = tlPercentage;
		this.ttPercentage = ttPercentage;
		this.noteOrderTl = noteOrderTl;
		this.status = status;
		this.creationDate = creationDate;
		this.createdBy = createdBy;
		this.lastUpdateDate = lastUpdateDate;
		this.lastUpdatedBy = lastUpdatedBy;
	}

	public BigDecimal getHeaderId() {
		return headerId;
	}

	public void setHeaderId(BigDecimal headerId) {
		this.headerId = headerId;
	}

	public String getMoId() {
		return moId;
	}

	public void setMoId(String moId) {
		this.moId = moId;
	}

	public Date getMonth() {
		return month;
	}

	public void setMonth(Date month) {
		this.month = month;
	}

	public BigDecimal getWdNormalTire() {
		return wdNormalTire;
	}

	public void setWdNormalTire(BigDecimal wdNormalTire) {
		this.wdNormalTire = wdNormalTire;
	}

	public BigDecimal getWdOtTl() {
		return wdOtTl;
	}

	public void setWdOtTl(BigDecimal wdOtTl) {
		this.wdOtTl = wdOtTl;
	}

	public BigDecimal getWdOtTt() {
		return wdOtTt;
	}

	public void setWdOtTt(BigDecimal wdOtTt) {
		this.wdOtTt = wdOtTt;
	}

	public BigDecimal getWdNormalTube() {
		return wdNormalTube;
	}

	public void setWdNormalTube(BigDecimal wdNormalTube) {
		this.wdNormalTube = wdNormalTube;
	}

	public BigDecimal getWdOtTube() {
		return wdOtTube;
	}

	public void setWdOtTube(BigDecimal wdOtTube) {
		this.wdOtTube = wdOtTube;
	}

	public BigDecimal getTotalWdTl() {
		return totalWdTl;
	}

	public void setTotalWdTl(BigDecimal totalWdTl) {
		this.totalWdTl = totalWdTl;
	}

	public BigDecimal getTotalWdTt() {
		return totalWdTt;
	}

	public void setTotalWdTt(BigDecimal totalWdTt) {
		this.totalWdTt = totalWdTt;
	}

	public BigDecimal getTotalWdTube() {
		return totalWdTube;
	}

	public void setTotalWdTube(BigDecimal totalWdTube) {
		this.totalWdTube = totalWdTube;
	}

	public BigDecimal getMaxCapTube() {
		return maxCapTube;
	}

	public void setMaxCapTube(BigDecimal maxCapTube) {
		this.maxCapTube = maxCapTube;
	}

	public BigDecimal getMaxCapTl() {
		return maxCapTl;
	}

	public void setMaxCapTl(BigDecimal maxCapTl) {
		this.maxCapTl = maxCapTl;
	}

	public BigDecimal getMaxCapTt() {
		return maxCapTt;
	}

	public void setMaxCapTt(BigDecimal maxCapTt) {
		this.maxCapTt = maxCapTt;
	}

	public BigDecimal getAirbagMachine() {
		return airbagMachine;
	}

	public void setAirbagMachine(BigDecimal airbagMachine) {
		this.airbagMachine = airbagMachine;
	}

	public BigDecimal getTl() {
		return tl;
	}

	public void setTl(BigDecimal tl) {
		this.tl = tl;
	}

	public BigDecimal getTt() {
		return tt;
	}

	public void setTt(BigDecimal tt) {
		this.tt = tt;
	}

	public BigDecimal getTotalMo() {
		return totalMo;
	}

	public void setTotalMo(BigDecimal totalMo) {
		this.totalMo = totalMo;
	}

	public BigDecimal getTlPercentage() {
		return tlPercentage;
	}

	public void setTlPercentage(BigDecimal tlPercentage) {
		this.tlPercentage = tlPercentage;
	}

	public BigDecimal getTtPercentage() {
		return ttPercentage;
	}

	public void setTtPercentage(BigDecimal ttPercentage) {
		this.ttPercentage = ttPercentage;
	}

	public String getNoteOrderTl() {
		return noteOrderTl;
	}

	public void setNoteOrderTl(String noteOrderTl) {
		this.noteOrderTl = noteOrderTl;
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
