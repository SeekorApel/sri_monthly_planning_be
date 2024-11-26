package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_T_MARKETINGORDER")
public class MarketingOrder {
	@Id
    @Column(name = "MO_ID")
    private String moId;
	
	@Column(name = "TYPE")
	private String type;
	
	@Column(name = "DATE_VALID")
	private Date dateValid;
	
	@Column(name = "REVISION_PPC")
	private BigDecimal revisionPpc;
	
	@Column(name = "REVISION_MARKETING")
	private BigDecimal revisionMarketing;
	
	@Column(name = "MONTH_0")
	private Date month0;
	
	@Column(name = "MONTH_1")
	private Date month1;
	
	@Column(name = "MONTH_2")
	private Date month2;
	
	@Column(name = "STATUS_FILLED")
	private BigDecimal statusFilled;
	
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
	
	public MarketingOrder() {}

	public MarketingOrder(MarketingOrder marketingOrder) {
        this.moId = marketingOrder.getMoId();
        this.type = marketingOrder.getType();
        this.dateValid = marketingOrder.getDateValid();
        this.revisionPpc = marketingOrder.getRevisionPpc();
        this.revisionMarketing = marketingOrder.getRevisionMarketing();
        this.month0 = marketingOrder.getMonth0();
        this.month1 = marketingOrder.getMonth1();
        this.month2 = marketingOrder.getMonth2();
        this.statusFilled = marketingOrder.getStatusFilled();
        this.status = marketingOrder.getStatus();
        this.creationDate = marketingOrder.getCreationDate();
        this.createdBy = marketingOrder.getCreatedBy();
        this.lastUpdateDate = marketingOrder.getLastUpdateDate();
        this.lastUpdatedBy = marketingOrder.getLastUpdatedBy();
    }


	public MarketingOrder(String moId, String type, Date dateValid, BigDecimal revisionPpc,
			BigDecimal revisionMarketing, Date month0, Date month1, Date month2, BigDecimal statusFilled,
			BigDecimal status, Date creationDate, String createdBy, Date lastUpdateDate, String lastUpdatedBy) {
		super();
		this.moId = moId;
		this.type = type;
		this.dateValid = dateValid;
		this.revisionPpc = revisionPpc;
		this.revisionMarketing = revisionMarketing;
		this.month0 = month0;
		this.month1 = month1;
		this.month2 = month2;
		this.statusFilled = statusFilled;
		this.status = status;
		this.creationDate = creationDate;
		this.createdBy = createdBy;
		this.lastUpdateDate = lastUpdateDate;
		this.lastUpdatedBy = lastUpdatedBy;
	}



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


	public BigDecimal getRevisionPpc() {
		return revisionPpc;
	}



	public void setRevisionPpc(BigDecimal revisionPpc) {
		this.revisionPpc = revisionPpc;
	}



	public BigDecimal getRevisionMarketing() {
		return revisionMarketing;
	}



	public void setRevisionMarketing(BigDecimal revisionMarketing) {
		this.revisionMarketing = revisionMarketing;
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
