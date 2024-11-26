package sri.sysint.sri_starter_back.model;

import java.math.BigDecimal;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "SRI_IMPP_D_WD_HOURS")
public class DWorkDayHours {
    
    @Id
    @Column(name = "DETAIL_WD_HOURS_ID")
    private BigDecimal DETAIL_WD_HOURS_ID;
    
    @Column(name = "DATE_WD")
    private Date DATE_WD;
    
    @Column(name = "HOUR_1")
    private BigDecimal HOUR_1;
    
    @Column(name = "HOUR_2")
    private BigDecimal HOUR_2;
    
    @Column(name = "HOUR_3")
    private BigDecimal HOUR_3;
    
    @Column(name = "HOUR_4")
    private BigDecimal HOUR_4;
    
    @Column(name = "HOUR_5")
    private BigDecimal HOUR_5;
    
    @Column(name = "HOUR_6")
    private BigDecimal HOUR_6;
    
    @Column(name = "HOUR_7")
    private BigDecimal HOUR_7;
    
    @Column(name = "HOUR_8")
    private BigDecimal HOUR_8;
    
    @Column(name = "HOUR_9")
    private BigDecimal HOUR_9;
    
    @Column(name = "HOUR_10")
    private BigDecimal HOUR_10;
    
    @Column(name = "HOUR_11")
    private BigDecimal HOUR_11;
    
    @Column(name = "HOUR_12")
    private BigDecimal HOUR_12;
    
    @Column(name = "HOUR_13")
    private BigDecimal HOUR_13;
    
    @Column(name = "HOUR_14")
    private BigDecimal HOUR_14;
    
    @Column(name = "HOUR_15")
    private BigDecimal HOUR_15;
    
    @Column(name = "HOUR_16")
    private BigDecimal HOUR_16;
    
    @Column(name = "HOUR_17")
    private BigDecimal HOUR_17;
    
    @Column(name = "HOUR_18")
    private BigDecimal HOUR_18;
    
    @Column(name = "HOUR_19")
    private BigDecimal HOUR_19;
    
    @Column(name = "HOUR_20")
    private BigDecimal HOUR_20;
    
    @Column(name = "HOUR_21")
    private BigDecimal HOUR_21;
    
    @Column(name = "HOUR_22")
    private BigDecimal HOUR_22;
    
    @Column(name = "HOUR_23")
    private BigDecimal HOUR_23;
    
    @Column(name = "HOUR_24")
    private BigDecimal HOUR_24;
    
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
    
    public DWorkDayHours() {
    	
    }
    
    public DWorkDayHours(DWorkDayHours workDay) {
    	this.DETAIL_WD_HOURS_ID = workDay.getDETAIL_WD_HOURS_ID();
        this.DATE_WD = workDay.getDATE_WD();
        this.HOUR_1 = workDay.getHOUR_1();
        this.HOUR_2 = workDay.getHOUR_2();
        this.HOUR_3 = workDay.getHOUR_3();
        this.HOUR_4 = workDay.getHOUR_4();
        this.HOUR_5 = workDay.getHOUR_5();
        this.HOUR_6 = workDay.getHOUR_6();
        this.HOUR_7 = workDay.getHOUR_7();
        this.HOUR_8 = workDay.getHOUR_8();
        this.HOUR_9 = workDay.getHOUR_9();
        this.HOUR_10 = workDay.getHOUR_10();
        this.HOUR_11 = workDay.getHOUR_11();
        this.HOUR_12 = workDay.getHOUR_12();
        this.HOUR_13 = workDay.getHOUR_13();
        this.HOUR_14 = workDay.getHOUR_14();
        this.HOUR_15 = workDay.getHOUR_15();
        this.HOUR_16 = workDay.getHOUR_16();
        this.HOUR_17 = workDay.getHOUR_17();
        this.HOUR_18 = workDay.getHOUR_18();
        this.HOUR_19 = workDay.getHOUR_19();
        this.HOUR_20 = workDay.getHOUR_20();
        this.HOUR_21 = workDay.getHOUR_21();
        this.HOUR_22 = workDay.getHOUR_22();
        this.HOUR_23 = workDay.getHOUR_23();
        this.HOUR_24 = workDay.getHOUR_24();
        this.DESCRIPTION = workDay.getDESCRIPTION();
        this.STATUS = workDay.getSTATUS();
        this.CREATION_DATE = workDay.getCREATION_DATE();
        this.CREATED_BY = workDay.getCREATED_BY();
        this.LAST_UPDATE_DATE = workDay.getLAST_UPDATE_DATE();
        this.LAST_UPDATED_BY = workDay.getLAST_UPDATED_BY();
    }

    
	public DWorkDayHours(BigDecimal dETAIL_WD_HOURS_ID, Date dATE_WD, BigDecimal hOUR_1, BigDecimal hOUR_2,
			BigDecimal hOUR_3, BigDecimal hOUR_4, BigDecimal hOUR_5, BigDecimal hOUR_6, BigDecimal hOUR_7,
			BigDecimal hOUR_8, BigDecimal hOUR_9, BigDecimal hOUR_10, BigDecimal hOUR_11, BigDecimal hOUR_12,
			BigDecimal hOUR_13, BigDecimal hOUR_14, BigDecimal hOUR_15, BigDecimal hOUR_16, BigDecimal hOUR_17,
			BigDecimal hOUR_18, BigDecimal hOUR_19, BigDecimal hOUR_20, BigDecimal hOUR_21, BigDecimal hOUR_22,
			BigDecimal hOUR_23, BigDecimal hOUR_24, String dESCRIPTION, BigDecimal sTATUS, Date cREATION_DATE,
			String cREATED_BY, Date lAST_UPDATE_DATE, String lAST_UPDATED_BY) {
		super();
		DETAIL_WD_HOURS_ID = dETAIL_WD_HOURS_ID;
		DATE_WD = dATE_WD;
		HOUR_1 = hOUR_1;
		HOUR_2 = hOUR_2;
		HOUR_3 = hOUR_3;
		HOUR_4 = hOUR_4;
		HOUR_5 = hOUR_5;
		HOUR_6 = hOUR_6;
		HOUR_7 = hOUR_7;
		HOUR_8 = hOUR_8;
		HOUR_9 = hOUR_9;
		HOUR_10 = hOUR_10;
		HOUR_11 = hOUR_11;
		HOUR_12 = hOUR_12;
		HOUR_13 = hOUR_13;
		HOUR_14 = hOUR_14;
		HOUR_15 = hOUR_15;
		HOUR_16 = hOUR_16;
		HOUR_17 = hOUR_17;
		HOUR_18 = hOUR_18;
		HOUR_19 = hOUR_19;
		HOUR_20 = hOUR_20;
		HOUR_21 = hOUR_21;
		HOUR_22 = hOUR_22;
		HOUR_23 = hOUR_23;
		HOUR_24 = hOUR_24;
		DESCRIPTION = dESCRIPTION;
		STATUS = sTATUS;
		CREATION_DATE = cREATION_DATE;
		CREATED_BY = cREATED_BY;
		LAST_UPDATE_DATE = lAST_UPDATE_DATE;
		LAST_UPDATED_BY = lAST_UPDATED_BY;
	}

	public BigDecimal getDETAIL_WD_HOURS_ID() {
		return DETAIL_WD_HOURS_ID;
	}

	public void setDETAIL_WD_HOURS_ID(BigDecimal dETAIL_WD_HOURS_ID) {
		DETAIL_WD_HOURS_ID = dETAIL_WD_HOURS_ID;
	}

	public Date getDATE_WD() {
		return DATE_WD;
	}

	public void setDATE_WD(Date dATE_WD) {
		DATE_WD = dATE_WD;
	}

	public BigDecimal getHOUR_1() {
		return HOUR_1;
	}

	public void setHOUR_1(BigDecimal hOUR_1) {
		HOUR_1 = hOUR_1;
	}

	public BigDecimal getHOUR_2() {
		return HOUR_2;
	}

	public void setHOUR_2(BigDecimal hOUR_2) {
		HOUR_2 = hOUR_2;
	}

	public BigDecimal getHOUR_3() {
		return HOUR_3;
	}

	public void setHOUR_3(BigDecimal hOUR_3) {
		HOUR_3 = hOUR_3;
	}

	public BigDecimal getHOUR_4() {
		return HOUR_4;
	}

	public void setHOUR_4(BigDecimal hOUR_4) {
		HOUR_4 = hOUR_4;
	}

	public BigDecimal getHOUR_5() {
		return HOUR_5;
	}

	public void setHOUR_5(BigDecimal hOUR_5) {
		HOUR_5 = hOUR_5;
	}

	public BigDecimal getHOUR_6() {
		return HOUR_6;
	}

	public void setHOUR_6(BigDecimal hOUR_6) {
		HOUR_6 = hOUR_6;
	}

	public BigDecimal getHOUR_7() {
		return HOUR_7;
	}

	public void setHOUR_7(BigDecimal hOUR_7) {
		HOUR_7 = hOUR_7;
	}

	public BigDecimal getHOUR_8() {
		return HOUR_8;
	}

	public void setHOUR_8(BigDecimal hOUR_8) {
		HOUR_8 = hOUR_8;
	}

	public BigDecimal getHOUR_9() {
		return HOUR_9;
	}

	public void setHOUR_9(BigDecimal hOUR_9) {
		HOUR_9 = hOUR_9;
	}

	public BigDecimal getHOUR_10() {
		return HOUR_10;
	}

	public void setHOUR_10(BigDecimal hOUR_10) {
		HOUR_10 = hOUR_10;
	}

	public BigDecimal getHOUR_11() {
		return HOUR_11;
	}

	public void setHOUR_11(BigDecimal hOUR_11) {
		HOUR_11 = hOUR_11;
	}

	public BigDecimal getHOUR_12() {
		return HOUR_12;
	}

	public void setHOUR_12(BigDecimal hOUR_12) {
		HOUR_12 = hOUR_12;
	}

	public BigDecimal getHOUR_13() {
		return HOUR_13;
	}

	public void setHOUR_13(BigDecimal hOUR_13) {
		HOUR_13 = hOUR_13;
	}

	public BigDecimal getHOUR_14() {
		return HOUR_14;
	}

	public void setHOUR_14(BigDecimal hOUR_14) {
		HOUR_14 = hOUR_14;
	}

	public BigDecimal getHOUR_15() {
		return HOUR_15;
	}

	public void setHOUR_15(BigDecimal hOUR_15) {
		HOUR_15 = hOUR_15;
	}

	public BigDecimal getHOUR_16() {
		return HOUR_16;
	}

	public void setHOUR_16(BigDecimal hOUR_16) {
		HOUR_16 = hOUR_16;
	}

	public BigDecimal getHOUR_17() {
		return HOUR_17;
	}

	public void setHOUR_17(BigDecimal hOUR_17) {
		HOUR_17 = hOUR_17;
	}

	public BigDecimal getHOUR_18() {
		return HOUR_18;
	}

	public void setHOUR_18(BigDecimal hOUR_18) {
		HOUR_18 = hOUR_18;
	}

	public BigDecimal getHOUR_19() {
		return HOUR_19;
	}

	public void setHOUR_19(BigDecimal hOUR_19) {
		HOUR_19 = hOUR_19;
	}

	public BigDecimal getHOUR_20() {
		return HOUR_20;
	}

	public void setHOUR_20(BigDecimal hOUR_20) {
		HOUR_20 = hOUR_20;
	}

	public BigDecimal getHOUR_21() {
		return HOUR_21;
	}

	public void setHOUR_21(BigDecimal hOUR_21) {
		HOUR_21 = hOUR_21;
	}

	public BigDecimal getHOUR_22() {
		return HOUR_22;
	}

	public void setHOUR_22(BigDecimal hOUR_22) {
		HOUR_22 = hOUR_22;
	}

	public BigDecimal getHOUR_23() {
		return HOUR_23;
	}

	public void setHOUR_23(BigDecimal hOUR_23) {
		HOUR_23 = hOUR_23;
	}

	public BigDecimal getHOUR_24() {
		return HOUR_24;
	}

	public void setHOUR_24(BigDecimal hOUR_24) {
		HOUR_24 = hOUR_24;
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
