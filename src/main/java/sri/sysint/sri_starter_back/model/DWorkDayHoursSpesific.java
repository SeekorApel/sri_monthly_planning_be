package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_D_WD_HOURS_SPECIFIC")
public class DWorkDayHoursSpesific {
    
    @Id
    @Column(name = "DETAIL_WD_HOURS_SPECIFIC_ID")
    private BigDecimal DETAIL_WD_HOURS_SPECIFIC_ID;
    
    @Column(name = "DATE_WD")
    private Date DATE_WD;
    
    
    @Column(name = "SHIFT1_START_TIME")
    private String SHIFT1_START_TIME;
    
    @Column(name = "SHIFT1_END_TIME")
    private String SHIFT1_END_TIME;
    
    @Column(name = "SHIFT1_TOTAL_TIME")
    private BigDecimal SHIFT1_TOTAL_TIME;
    
    @Column(name = "SHIFT2_START_TIME")
    private String SHIFT2_START_TIME;
    
    @Column(name = "SHIFT2_END_TIME")
    private String SHIFT2_END_TIME;
    
    @Column(name = "SHIFT2_TOTAL_TIME")
    private BigDecimal SHIFT2_TOTAL_TIME;
    
    @Column(name = "SHIFT3_START_TIME")
    private String SHIFT3_START_TIME;
    
    @Column(name = "SHIFT3_END_TIME")
    private String SHIFT3_END_TIME;
    
    @Column(name = "SHIFT3_TOTAL_TIME")
    private BigDecimal SHIFT3_TOTAL_TIME;
    
    @Column(name = "DESCRIPTION")
    private String DESCRIPTION;
    
    @Column(name = "STATUS")
    private BigDecimal STATUS;
    
    @Column(name = "CREATION_DATE")
    private Date CREATION_DATE;
    
    @Column(name = "CREATED_BY")
    private String CREATED_BY;
    
    @Column(name = "LAST_UPDATE_DATE")
    private Date LAST_UPDATE_DATE;
    
    @Column(name = "LAST_UPDATED_BY")
    private String LAST_UPDATED_BY;
    
    public DWorkDayHoursSpesific() {
    	
    }
    
    public DWorkDayHoursSpesific(DWorkDayHoursSpesific dWorkDayHoursSpesific) {
        this.DETAIL_WD_HOURS_SPECIFIC_ID = dWorkDayHoursSpesific.getDETAIL_WD_HOURS_SPECIFIC_ID();
        this.DATE_WD = dWorkDayHoursSpesific.getDATE_WD();
        this.SHIFT1_START_TIME = dWorkDayHoursSpesific.getSHIFT1_START_TIME();
        this.SHIFT1_END_TIME = dWorkDayHoursSpesific.getSHIFT1_END_TIME();
        this.SHIFT1_TOTAL_TIME = dWorkDayHoursSpesific.getSHIFT1_TOTAL_TIME();
        this.SHIFT2_START_TIME = dWorkDayHoursSpesific.getSHIFT2_START_TIME();
        this.SHIFT2_END_TIME = dWorkDayHoursSpesific.getSHIFT2_END_TIME();
        this.SHIFT2_TOTAL_TIME = dWorkDayHoursSpesific.getSHIFT2_TOTAL_TIME();
        this.SHIFT3_START_TIME = dWorkDayHoursSpesific.getSHIFT3_START_TIME();
        this.SHIFT3_END_TIME = dWorkDayHoursSpesific.getSHIFT3_END_TIME();
        this.SHIFT3_TOTAL_TIME = dWorkDayHoursSpesific.getSHIFT3_TOTAL_TIME();
        this.DESCRIPTION = dWorkDayHoursSpesific.getDESCRIPTION();
        this.STATUS = dWorkDayHoursSpesific.getSTATUS();
        this.CREATION_DATE = dWorkDayHoursSpesific.getCREATION_DATE();
        this.CREATED_BY = dWorkDayHoursSpesific.getCREATED_BY();
        this.LAST_UPDATE_DATE = dWorkDayHoursSpesific.getLAST_UPDATE_DATE();
        this.LAST_UPDATED_BY = dWorkDayHoursSpesific.getLAST_UPDATED_BY();
    }
    
    

	public DWorkDayHoursSpesific(BigDecimal dETAIL_WD_HOURS_SPECIFIC_ID, Date dATE_WD, String sHIFT1_START_TIME,
			String sHIFT1_END_TIME, BigDecimal sHIFT1_TOTAL_TIME, String sHIFT2_START_TIME, String sHIFT2_END_TIME,
			BigDecimal sHIFT2_TOTAL_TIME, String sHIFT3_START_TIME, String sHIFT3_END_TIME,
			BigDecimal sHIFT3_TOTAL_TIME, String dESCRIPTION, BigDecimal sTATUS, Date cREATION_DATE, String cREATED_BY,
			Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		DETAIL_WD_HOURS_SPECIFIC_ID = dETAIL_WD_HOURS_SPECIFIC_ID;
		DATE_WD = dATE_WD;
		SHIFT1_START_TIME = sHIFT1_START_TIME;
		SHIFT1_END_TIME = sHIFT1_END_TIME;
		SHIFT1_TOTAL_TIME = sHIFT1_TOTAL_TIME;
		SHIFT2_START_TIME = sHIFT2_START_TIME;
		SHIFT2_END_TIME = sHIFT2_END_TIME;
		SHIFT2_TOTAL_TIME = sHIFT2_TOTAL_TIME;
		SHIFT3_START_TIME = sHIFT3_START_TIME;
		SHIFT3_END_TIME = sHIFT3_END_TIME;
		SHIFT3_TOTAL_TIME = sHIFT3_TOTAL_TIME;
		DESCRIPTION = dESCRIPTION;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getDETAIL_WD_HOURS_SPECIFIC_ID() {
		return DETAIL_WD_HOURS_SPECIFIC_ID;
	}

	public void setDETAIL_WD_HOURS_SPECIFIC_ID(BigDecimal dETAIL_WD_HOURS_SPECIFIC_ID) {
		DETAIL_WD_HOURS_SPECIFIC_ID = dETAIL_WD_HOURS_SPECIFIC_ID;
	}

	public Date getDATE_WD() {
		return DATE_WD;
	}

	public void setDATE_WD(Date dATE_WD) {
		DATE_WD = dATE_WD;
	}

	public String getSHIFT1_START_TIME() {
		return SHIFT1_START_TIME;
	}

	public void setSHIFT1_START_TIME(String sHIFT1_START_TIME) {
		SHIFT1_START_TIME = sHIFT1_START_TIME;
	}

	public String getSHIFT1_END_TIME() {
		return SHIFT1_END_TIME;
	}

	public void setSHIFT1_END_TIME(String sHIFT1_END_TIME) {
		SHIFT1_END_TIME = sHIFT1_END_TIME;
	}

	public BigDecimal getSHIFT1_TOTAL_TIME() {
		return SHIFT1_TOTAL_TIME;
	}

	public void setSHIFT1_TOTAL_TIME(BigDecimal sHIFT1_TOTAL_TIME) {
		SHIFT1_TOTAL_TIME = sHIFT1_TOTAL_TIME;
	}

	public String getSHIFT2_START_TIME() {
		return SHIFT2_START_TIME;
	}

	public void setSHIFT2_START_TIME(String sHIFT2_START_TIME) {
		SHIFT2_START_TIME = sHIFT2_START_TIME;
	}

	public String getSHIFT2_END_TIME() {
		return SHIFT2_END_TIME;
	}

	public void setSHIFT2_END_TIME(String sHIFT2_END_TIME) {
		SHIFT2_END_TIME = sHIFT2_END_TIME;
	}

	public BigDecimal getSHIFT2_TOTAL_TIME() {
		return SHIFT2_TOTAL_TIME;
	}

	public void setSHIFT2_TOTAL_TIME(BigDecimal sHIFT2_TOTAL_TIME) {
		SHIFT2_TOTAL_TIME = sHIFT2_TOTAL_TIME;
	}

	public String getSHIFT3_START_TIME() {
		return SHIFT3_START_TIME;
	}

	public void setSHIFT3_START_TIME(String sHIFT3_START_TIME) {
		SHIFT3_START_TIME = sHIFT3_START_TIME;
	}

	public String getSHIFT3_END_TIME() {
		return SHIFT3_END_TIME;
	}

	public void setSHIFT3_END_TIME(String sHIFT3_END_TIME) {
		SHIFT3_END_TIME = sHIFT3_END_TIME;
	}

	public BigDecimal getSHIFT3_TOTAL_TIME() {
		return SHIFT3_TOTAL_TIME;
	}

	public void setSHIFT3_TOTAL_TIME(BigDecimal sHIFT3_TOTAL_TIME) {
		SHIFT3_TOTAL_TIME = sHIFT3_TOTAL_TIME;
	}

	public String getDESCRIPTION() {
		return DESCRIPTION;
	}

	public void setDESCRIPTION(String dESCRIPTION) {
		DESCRIPTION = dESCRIPTION;
	}

	public BigDecimal getSTATUS() {
		return STATUS;
	}

	public void setSTATUS(BigDecimal sTATUS) {
		STATUS = sTATUS;
	}

	public Date getCREATION_DATE() {
		return CREATION_DATE;
	}

	public void setCREATION_DATE(Date cREATION_DATE) {
		CREATION_DATE = cREATION_DATE;
	}

	public String getCREATED_BY() {
		return CREATED_BY;
	}

	public void setCREATED_BY(String cREATED_BY) {
		CREATED_BY = cREATED_BY;
	}

	public Date getLAST_UPDATE_DATE() {
		return LAST_UPDATE_DATE;
	}

	public void setLAST_UPDATE_DATE(Date lAST_UPDATE_DATE) {
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
	}

	public String getLAST_UPDATED_BY() {
		return LAST_UPDATED_BY;
	}

	public void setLAST_UPDATED_BY(String lAST_UPDATED_BY) {
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

   
	
    
}
