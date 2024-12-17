package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_D_MONTHLYPLAN_CURING")
public class DetailMonthlyPlanCuring {
	
	@Id
    @Column(name = "DETAIL_ID_CURING")
    private BigDecimal detailIdCuring;
	
	@Column(name = "DOC_NUM_CURING")
	private String docNumber;
	
    @Column(name = "PART_NUMBER")
    private BigDecimal partNumber;
	
	@Column(name = "TOTAL")
	private BigDecimal total;
	
    @Column(name = "NET_FULFILMENT")
    private BigDecimal netFulfilment;
	
	@Column(name = "GROSS_REQ")
	private BigDecimal grossReq;
	
	@Column(name = "NET_REQ")
    private BigDecimal netReq;
	
	@Column(name = "REQ_AHM_OEM")
    private BigDecimal reqAhmOem;
	
	@Column(name = "REQ_AHM_REM")
    private BigDecimal reqAhmRem;
	
	@Column(name = "REQ_FDR")
	private BigDecimal reqFdr;
	
	@Column(name = "DIFFERENCE_OFS")
    private BigDecimal differenceOfs;
	
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
	
	

	public DetailMonthlyPlanCuring() {
	}

	
	
	public DetailMonthlyPlanCuring(BigDecimal detailIdCuring, String docNumber, BigDecimal partNumber, BigDecimal total,
			BigDecimal netFulfilment, BigDecimal grossReg, BigDecimal netReq, BigDecimal reqAhmOem,
			BigDecimal reqAhmRem, BigDecimal reqFdr, BigDecimal differenceOfs, BigDecimal status, Date creationDate,
			String createdBy, Date lastUpdateDate, String lastUpdatedBy) {
		this.detailIdCuring = detailIdCuring;
		this.docNumber = docNumber;
		this.partNumber = partNumber;
		this.total = total;
		this.netFulfilment = netFulfilment;
		this.grossReq = grossReq;
		this.netReq = netReq;
		this.reqAhmOem = reqAhmOem;
		this.reqAhmRem = reqAhmRem;
		this.reqFdr = reqFdr;
		this.differenceOfs = differenceOfs;
		this.status = status;
		this.creationDate = creationDate;
		this.createdBy = createdBy;
		this.lastUpdateDate = lastUpdateDate;
		this.lastUpdatedBy = lastUpdatedBy;
	}
	
	public DetailMonthlyPlanCuring(DetailMonthlyPlanCuring detailMonthlyPlanCuring) {
	    super();
	    this.detailIdCuring = detailMonthlyPlanCuring.getDetailIdCuring();
	    this.docNumber = detailMonthlyPlanCuring.getDocNumber();
	    this.partNumber = detailMonthlyPlanCuring.getPartNumber();
	    this.total = detailMonthlyPlanCuring.getTotal();
	    this.netFulfilment = detailMonthlyPlanCuring.getNetFulfilment();
	    this.grossReq = detailMonthlyPlanCuring.getGrossReq();
	    this.netReq = detailMonthlyPlanCuring.getNetReq();
	    this.reqAhmOem = detailMonthlyPlanCuring.getReqAhmOem();
	    this.reqAhmRem = detailMonthlyPlanCuring.getReqAhmRem();
	    this.reqFdr = detailMonthlyPlanCuring.getReqFdr();
	    this.differenceOfs = detailMonthlyPlanCuring.getDifferenceOfs();
	    this.status = detailMonthlyPlanCuring.getStatus();
	    this.creationDate = detailMonthlyPlanCuring.getCreationDate();
	    this.createdBy = detailMonthlyPlanCuring.getCreatedBy();
	    this.lastUpdateDate = detailMonthlyPlanCuring.getLastUpdateDate();
	    this.lastUpdatedBy = detailMonthlyPlanCuring.getLastUpdatedBy();
	}


	public BigDecimal getDetailIdCuring() {
		return detailIdCuring;
	}

	public void setDetailIdCuring(BigDecimal detailIdCuring) {
		this.detailIdCuring = detailIdCuring;
	}

	public String getDocNumber() {
		return docNumber;
	}

	public void setDocNumber(String docNumber) {
		this.docNumber = docNumber;
	}

	public BigDecimal getPartNumber() {
		return partNumber;
	}
	

	public void setPartNumber(BigDecimal partNumber) {
		this.partNumber = partNumber;
	}

	public BigDecimal getNetReq() {
		return netReq;
	}

	public void setNetReq(BigDecimal netReq) {
		this.netReq = netReq;
	}

	public BigDecimal getReqAhmOem() {
		return reqAhmOem;
	}

	public void setReqAhmOem(BigDecimal reqAhmOem) {
		this.reqAhmOem = reqAhmOem;
	}

	public BigDecimal getTotal() {
		return total;
	}

	public void setTotal(BigDecimal total) {
		this.total = total;
	}

	public BigDecimal getNetFulfilment() {
		return netFulfilment;
	}

	public void setNetFulfilment(BigDecimal netFulfilment) {
		this.netFulfilment = netFulfilment;
	}



	public BigDecimal getGrossReq() {
		return grossReq;
	}



	public void setGrossReq(BigDecimal grossReq) {
		this.grossReq = grossReq;
	}



	public BigDecimal getReqAhmRem() {
		return reqAhmRem;
	}

	public void setReqAhmRem(BigDecimal reqAhmRem) {
		this.reqAhmRem = reqAhmRem;
	}

	public BigDecimal getReqFdr() {
		return reqFdr;
	}

	public void setReqFdr(BigDecimal reqFdr) {
		this.reqFdr = reqFdr;
	}

	public BigDecimal getDifferenceOfs() {
		return differenceOfs;
	}

	public void setDifferenceOfs(BigDecimal differenceOfs) {
		this.differenceOfs = differenceOfs;
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
