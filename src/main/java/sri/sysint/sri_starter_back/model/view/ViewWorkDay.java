
package sri.sysint.sri_starter_back.model.view;

import java.math.BigDecimal;
import java.util.Date;

import sri.sysint.sri_starter_back.model.WorkDay;

public class ViewWorkDay {
	private Date dateWd;
    private BigDecimal iwdShift1;
    private BigDecimal iwdShift2;
    private BigDecimal iwdShift3;
    private BigDecimal iotTl1;
    private BigDecimal iotTl2;
    private BigDecimal iotTl3;
    private BigDecimal iotTt1;
    private BigDecimal iotTt2;
    private BigDecimal iotTt3;
    private BigDecimal off;
    
    
	    
    public ViewWorkDay(WorkDay workDay) {
        this.dateWd = workDay.getDATE_WD();
        this.iwdShift1 = workDay.getIWD_SHIFT_1();
        this.iwdShift2 = workDay.getIWD_SHIFT_2();
        this.iwdShift3 = workDay.getIWD_SHIFT_3();
        this.iotTl1 = workDay.getIOT_TL_1();
        this.iotTl2 = workDay.getIOT_TL_2();
        this.iotTl3 = workDay.getIOT_TL_3();
        this.iotTt1 = workDay.getIOT_TT_1();
        this.iotTt2 = workDay.getIOT_TT_2();
        this.iotTt3 = workDay.getIOT_TT_3();
        this.off = workDay.getOFF();
    }


		public Date getDateWd() {
			return dateWd;
		}


		public void setDateWd(Date dateWd) {
			this.dateWd = dateWd;
		}


		public BigDecimal getIwdShift1() {
			return iwdShift1;
		}


		public void setIwdShift1(BigDecimal iwdShift1) {
			this.iwdShift1 = iwdShift1;
		}


		public BigDecimal getIwdShift2() {
			return iwdShift2;
		}


		public void setIwdShift2(BigDecimal iwdShift2) {
			this.iwdShift2 = iwdShift2;
		}


		public BigDecimal getIwdShift3() {
			return iwdShift3;
		}


		public void setIwdShift3(BigDecimal iwdShift3) {
			this.iwdShift3 = iwdShift3;
		}


		public BigDecimal getIotTl1() {
			return iotTl1;
		}


		public void setIotTl1(BigDecimal iotTl1) {
			this.iotTl1 = iotTl1;
		}


		public BigDecimal getIotTl2() {
			return iotTl2;
		}


		public void setIotTl2(BigDecimal iotTl2) {
			this.iotTl2 = iotTl2;
		}


		public BigDecimal getIotTl3() {
			return iotTl3;
		}


		public void setIotTl3(BigDecimal iotTl3) {
			this.iotTl3 = iotTl3;
		}


		public BigDecimal getIotTt1() {
			return iotTt1;
		}


		public void setIotTt1(BigDecimal iotTt1) {
			this.iotTt1 = iotTt1;
		}


		public BigDecimal getIotTt2() {
			return iotTt2;
		}


		public void setIotTt2(BigDecimal iotTt2) {
			this.iotTt2 = iotTt2;
		}


		public BigDecimal getIotTt3() {
			return iotTt3;
		}


		public void setIotTt3(BigDecimal iotTt3) {
			this.iotTt3 = iotTt3;
		}


		public BigDecimal getOff() {
			return off;
		}


		public void setOff(BigDecimal off) {
			this.off = off;
		}
		
		
}
