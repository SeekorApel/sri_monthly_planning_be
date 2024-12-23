package sri.sysint.sri_starter_back.model;
import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;


@Entity
@Table(name = "SRI_IMPP_CHANGE_MOULD")
public class ChangeMould {
	@Id
	@Column(name = "ID_CHANGE_MOULD")
	private BigDecimal idChangeNumber;

	@Column(name = "PART_NUMBER")
	private BigDecimal partNum;
	
	@Column(name = "CHANGE_DATE")
	private Date changeDate;

	@Column(name = "WORK_CENTER_TEXT")
	private String wct;
	
	@Column(name = "SHIFT")
	private BigDecimal shift;

	public ChangeMould() {}
	
	public ChangeMould(BigDecimal idChangeNumber, BigDecimal partNum, Date changeDate, String wct, BigDecimal shift) {
		super();
		this.idChangeNumber = idChangeNumber;
		this.partNum = partNum;
		this.changeDate = changeDate;
		this.wct = wct;
		this.shift = shift;
	}

	public BigDecimal getIdChangeNumber() {
		return idChangeNumber;
	}

	public void setIdChangeNumber(BigDecimal idChangeNumber) {
		this.idChangeNumber = idChangeNumber;
	}

	public BigDecimal getPartNum() {
		return partNum;
	}

	public void setPartNum(BigDecimal partNum) {
		this.partNum = partNum;
	}

	public Date getChangeDate() {
		return changeDate;
	}

	public void setChangeDate(Date changeDate) {
		this.changeDate = changeDate;
	}

	public String getWct() {
		return wct;
	}

	public void setWct(String wct) {
		this.wct = wct;
	}

	public BigDecimal getShift() {
		return shift;
	}

	public void setShift(BigDecimal shift) {
		this.shift = shift;
	}
	
	
}
