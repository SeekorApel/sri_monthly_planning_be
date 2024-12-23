package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_T_MONTHLYPLAN_CURING")
public class MonthlyPlanningCuring {
	
	@Id
    @Column(name = "MP_CURING_ID")
    private String mpCuringId;
	
	@Column(name = "DOC_NUMBER")
	private String docNumber;
	
	@Column(name = "EFFECTIVE_DATE")
	private Date effectiveDate;
	
	@Column(name = "REVISION")
	private BigDecimal revision;
	
	@Column(name = "WORK_DAY")
	private BigDecimal workDay;
	
	@Column(name = "OVERTIME")
	private BigDecimal overtime;
	
	@Column(name = "MONTH_OF")
	private Date monthOf;
	
    @Column(name = "KADEPT")
    private String kadept;
	
	@Column(name = "KASSIE_PP")
	private String kassiePp;
	
	@Column(name = "SECTION")
	private String section;
	
	@Column(name = "ISSUE_DATE")
	private Date issueDate;
	
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

	public MonthlyPlanningCuring(String mpCuringId, String docNumCuring, Date effectiveDate, BigDecimal revision,
			BigDecimal workDat, BigDecimal overtime, Date monthOf, String kadept, String kassiePp, String section,
			Date issueDate, BigDecimal status, Date creationDate, String createdBy, Date lastUpdateDate,
			String lastUpdatedBy) {
		super();
		this.mpCuringId = mpCuringId;
		this.docNumber = docNumCuring;
		this.effectiveDate = effectiveDate;
		this.revision = revision;
		this.workDay = workDat;
		this.overtime = overtime;
		this.monthOf = monthOf;
		this.kadept = kadept;
		this.kassiePp = kassiePp;
		this.section = section;
		this.issueDate = issueDate;
		this.status = status;
		this.creationDate = creationDate;
		this.createdBy = createdBy;
		this.lastUpdateDate = lastUpdateDate;
		this.lastUpdatedBy = lastUpdatedBy;
	}
	
	public MonthlyPlanningCuring() {}

	public MonthlyPlanningCuring(MonthlyPlanningCuring monthlyPlanningCuring) {
	    this.mpCuringId = monthlyPlanningCuring.getMpCuringId();
	    this.docNumber = monthlyPlanningCuring.getDocNumber();
	    this.effectiveDate = monthlyPlanningCuring.getEffectiveDate();
	    this.revision = monthlyPlanningCuring.getRevision();
	    this.workDay = monthlyPlanningCuring.getWorkDay();
	    this.overtime = monthlyPlanningCuring.getOvertime();
	    this.monthOf = monthlyPlanningCuring.getMonthOf();
	    this.kadept = monthlyPlanningCuring.getKadept();
	    this.kassiePp = monthlyPlanningCuring.getKassiePp();
	    this.section = monthlyPlanningCuring.getSection();
	    this.issueDate = monthlyPlanningCuring.getIssueDate();
	    this.status = monthlyPlanningCuring.getStatus();
	    this.creationDate = monthlyPlanningCuring.getCreationDate();
	    this.createdBy = monthlyPlanningCuring.getCreatedBy();
	    this.lastUpdateDate = monthlyPlanningCuring.getLastUpdateDate();
	    this.lastUpdatedBy = monthlyPlanningCuring.getLastUpdatedBy();
	}


	public String getMpCuringId() {
		return mpCuringId;
	}

	public void setMpCuringId(String mpCuringId) {
		this.mpCuringId = mpCuringId;
	}

	

	public Date getEffectiveDate() {
		return effectiveDate;
	}

	public void setEffectiveDate(Date effectiveDate) {
		this.effectiveDate = effectiveDate;
	}

	public BigDecimal getRevision() {
		return revision;
	}

	public void setRevision(BigDecimal revision) {
		this.revision = revision;
	}

	
	public String getDocNumber() {
		return docNumber;
	}

	public void setDocNumber(String docNumber) {
		this.docNumber = docNumber;
	}

	public BigDecimal getWorkDay() {
		return workDay;
	}

	public void setWorkDay(BigDecimal workDay) {
		this.workDay = workDay;
	}

	public BigDecimal getOvertime() {
		return overtime;
	}

	public void setOvertime(BigDecimal overtime) {
		this.overtime = overtime;
	}

	public Date getMonthOf() {
		return monthOf;
	}

	public void setMonthOf(Date monthOf) {
		this.monthOf = monthOf;
	}

	public String getKadept() {
		return kadept;
	}

	public void setKadept(String kadept) {
		this.kadept = kadept;
	}

	public String getKassiePp() {
		return kassiePp;
	}

	public void setKassiePp(String kassiePp) {
		this.kassiePp = kassiePp;
	}

	public String getSection() {
		return section;
	}

	public void setSection(String section) {
		this.section = section;
	}

	public Date getIssueDate() {
		return issueDate;
	}

	public void setIssueDate(Date issueDate) {
		this.issueDate = issueDate;
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
