package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;


@Entity
@Table(name = "SRI_IMPP_M_MACHINE_PRODUCT")
public class MachineProduct {
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "PART_NUMBER", nullable = false)
    private BigDecimal PART_NUMBER;

	@Column(name = "WORK_CENTER_TEXT", nullable = false)
	private String WORK_CENTER_TEXT;
	

	public MachineProduct() {
		super();
	}

	public MachineProduct(BigDecimal pART_NUMBER, String wORK_CENTER_TEXT) {
		super();
		PART_NUMBER = pART_NUMBER;
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
	}

	public BigDecimal getPART_NUMBER() {
		return PART_NUMBER;
	}

	public void setPART_NUMBER(BigDecimal pART_NUMBER) {
		PART_NUMBER = pART_NUMBER;
	}

	public String getWORK_CENTER_TEXT() {
		return WORK_CENTER_TEXT;
	}

	public void setWORK_CENTER_TEXT(String wORK_CENTER_TEXT) {
		WORK_CENTER_TEXT = wORK_CENTER_TEXT;
	}
}
