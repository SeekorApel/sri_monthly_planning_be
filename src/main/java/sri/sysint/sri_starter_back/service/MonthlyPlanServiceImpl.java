package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.Iterator;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.CTCuring;
import sri.sysint.sri_starter_back.model.DetailDailyMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.DetailMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineProduct;
import sri.sysint.sri_starter_back.model.MarketingOrder;
import sri.sysint.sri_starter_back.model.ShiftMonthlyPlan;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.model.transaksi.ViewMonthlyPlanning;
import sri.sysint.sri_starter_back.repository.CTCuringRepo;
import sri.sysint.sri_starter_back.repository.DWorkDayHoursSpecificRepo;
import sri.sysint.sri_starter_back.repository.DetailMarketingOrderRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;
import sri.sysint.sri_starter_back.repository.MachineProductRepo;
import sri.sysint.sri_starter_back.repository.MarketingOrderRepo;
import sri.sysint.sri_starter_back.repository.MonthlyPlanRepo;
import sri.sysint.sri_starter_back.repository.SettingRepo;
import sri.sysint.sri_starter_back.repository.ShiftMonthlyPlanRepo;
import sri.sysint.sri_starter_back.repository.WorkDayRepo;


@Service
@Transactional
public class MonthlyPlanServiceImpl {
	
	@Autowired
    private MarketingOrderRepo marketingOrderRepo;
	
	@Autowired
    private MonthlyPlanRepo monthlyPlanningRepo;
	
	@Autowired
    private DetailMarketingOrderRepo detailMarketingOrderRepo;
	
	@Autowired
    private CTCuringRepo ctCuringRepo;
	
	@Autowired
    private ShiftMonthlyPlanRepo shiftMonthlyRepo;
	
	@Autowired
    private MachineCuringRepo machineCuringRepo;
	
	@Autowired
    private WorkDayRepo workDayRepo;
	
	@Autowired
    private SettingRepo settingRepo;
	
	@Autowired 
	private MachineProductRepo machineProductRepo;
	
	@Autowired
    private DWorkDayHoursSpecificRepo dWorkDayHourSpecificRepo;
	
    public MonthlyPlanServiceImpl(MonthlyPlanRepo monthlyPlanRepo){
    	
    }
    
    private List<MachineCuring> machineCuringList = new ArrayList<>();

	private List<MachineCuring> machineCuringListTemp = new ArrayList<>();
	
	private List<ShiftMonthlyPlan> newShiftMonthlyPlan = new ArrayList<>();
	
	private List<ShiftMonthlyPlan> newShiftListPartNum = new ArrayList<>();
	
	private List<WorkDay> workDayList = new ArrayList<>();
	
	private List<ShiftMonthlyPlan> newShiftList = new ArrayList<>();
	
	private List<Map<String, Object>> detailMarkOrderList = new ArrayList<>();
	
	private BigDecimal smallOrderLimit;
	
	private List<ShiftMonthlyPlan> oldShiftPlan = new ArrayList<>();
	
	private List<TempOrder> tempOrderList = new ArrayList<>();
	
	private List<MachineProduct> machineProductList = new ArrayList<>();
	
	private List<MachineCuring> machineCuringOrderList = new ArrayList<>();
	
	private List<MachineCuring> machineCuringUsedList = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderList = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderListAB = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderListBOM = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderListDual = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderListABFrontRear = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderListBOMFrontRear = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderListDualFrontRear = new ArrayList<>();
	
	private List<DetailMo> detailMarketingOrderFrontRear = new ArrayList<>();
	
	private List<ChangeMould> changeMouldList = new ArrayList<>();
	
	private List<ChangeMould> endProductList = new ArrayList<>();
	
	private List<TempChangeMould> tempChangeMouldList = new ArrayList<>();
	
	private BigDecimal order;
	
	private BigDecimal percentagePlusMinus;
	
	private int maxChangeMould;
	
	private BigDecimal minProduction;
	
	public List<ShiftMonthlyPlan> MonthlyPlan(int month, int year, int limitChange, BigDecimal minA, BigDecimal maxA, BigDecimal minB, BigDecimal maxB, BigDecimal minC, BigDecimal maxC, BigDecimal minD, BigDecimal maxD) {
//		percentagePlusMinus = percentage;
		maxChangeMould = limitChange;
		getDataHeader(month, year);
        boolean tempShift = false;
		if(machineProductList != null) {
			for(MachineProduct mn : machineProductList) {
	        	for(DetailMo dtMo : detailMarketingOrderListAB) {
	        		if(dtMo.getPartNumber().equals(mn.getPART_NUMBER())) {
	        			minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	        			order = dtMo.getTotalAR();
	        			tempShift = generateFromManualMapping(mn.getPART_NUMBER(), month, year, dtMo.getItemCuring(), mn.getWORK_CENTER_TEXT());
	        			if(tempShift == true) {
	        				dtMo.setMoMonth0(order);
	        			}
	        		}
	        	}
	        	for(DetailMo dtMo : detailMarketingOrderListDual) {
	        		if(dtMo.getPartNumber().equals(mn.getPART_NUMBER())) {
	        			order = dtMo.getTotalAR();
	        			minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	        			tempShift = generateFromManualMapping(mn.getPART_NUMBER(), month, year, dtMo.getItemCuring(), mn.getWORK_CENTER_TEXT());
	        			if(tempShift == true) {
	        				dtMo.setMoMonth0(order);
	        			}
	        		}
	        	}
	        	for(DetailMo dtMo : detailMarketingOrderListBOM) {
	        		if(dtMo.getPartNumber().equals(mn.getPART_NUMBER())) {
	        			order = dtMo.getTotalAR();
	        			minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	        			tempShift = generateFromManualMapping(mn.getPART_NUMBER(), month, year, dtMo.getItemCuring(), mn.getWORK_CENTER_TEXT());
	        			if(tempShift == true) {
	        				dtMo.setMoMonth0(order);
	        			}
	        		}
	        	}
	        }
		}
		
		for(DetailMo dtMo : detailMarketingOrderListABFrontRear) {
		////System.out.println("test 2");
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 3");
		    order = dtMo.getTotalAR();
		   ////System.out.println("test 4");
    	    oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), dtMo.getItemCuring());
    	   ////System.out.println("test 5");
    	    if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) { // if 1
    			minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
    	    	while (order.intValue() > minProduction.intValue()) {
        	    ////System.out.println("test 6");
        	    	if(!checkAllActiveMachine()) {
        	    		break;
        	    	}
        	    ////System.out.println("test 7");

    	            tempShift = generateFromOldShift(month, year);
    	            
        	       ////System.out.println("test 8");
        	        if(!tempShift){
        	        	break;
        	        }
        	    }
        	   ////System.out.println("test 9");
        	    dtMo.setMoMonth0(order);
	        }
    	    machineCuringList = machineCuringListTemp;
		}
		
		for(DetailMo dtMo : detailMarketingOrderListBOMFrontRear) {
			////System.out.println("test 2");
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 3");
		    order = dtMo.getTotalAR();
		   ////System.out.println("test 4");
    	    oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), dtMo.getItemCuring());
    	   ////System.out.println("test 5");
    	    if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) { // if 1
    	    	minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
    	    	while (order.intValue() > minProduction.intValue()) {
        	    ////System.out.println("test 6");
        	    	if(!checkAllActiveMachine()) {
        	    		break;
        	    	}
        	    ////System.out.println("test 7");

    	            tempShift = generateFromOldShift(month, year);
    	            
        	       ////System.out.println("test 8");
        	        if(!tempShift){
        	        	break;
        	        }
        	    }
        	   ////System.out.println("test 9");
        	    dtMo.setMoMonth0(order);
	        }
    	    machineCuringList = machineCuringListTemp;
		}
		
		for(DetailMo dtMo : detailMarketingOrderListDualFrontRear) {
		////System.out.println("test 2");
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 3");
		    order = dtMo.getTotalAR();
		   ////System.out.println("test 4");
    	    oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), dtMo.getItemCuring());
    	   ////System.out.println("test 5");
    	    if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) { // if 1
    	    	minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
    	    	while (order.intValue() > minProduction.intValue()) {
        	    ////System.out.println("test 6");
        	    	if(!checkAllActiveMachine()) {
        	    		break;
        	    	}
        	    	////System.out.println("test 7");

    	            tempShift = generateFromOldShift(month, year);
    	            
        	       ////System.out.println("test 8");
        	        if(!tempShift){
        	        	break;
        	        }
        	    }
        	   ////System.out.println("test 9");
        	    dtMo.setMoMonth0(order);
	        }
    	    machineCuringList = machineCuringListTemp;
		}
		
		for(DetailMo dtMo : detailMarketingOrderListABFrontRear) {
		////System.out.println("test 11");
			if(!checkAllActiveMachine()) {
				break;
			}
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 12");
            order = dtMo.getTotalAR();
           ////System.out.println("test 13");
            minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	    	while (order.intValue() > minProduction.intValue()) {
            ////System.out.println("test 14");
            	List<CTCuring> ctCurList = getMachine(statusPrioritasMesin, dtMo.getItemCuring());
	            if(ctCurList.size() > 0) {
	            ////System.out.println("test 15");
    	            tempShift = generateFromMidMonth(dtMo.getPartNumber(), month, year,  ctCurList);
	            }
	           ////System.out.println("test 16");
	            if(!tempShift) {
	            ////System.out.println("test 17");
	            	clearShift(dtMo.getPartNumber());
	            ////System.out.println("test 18");
	            	if (statusPrioritasMesin == 1) {
    	        		statusPrioritasMesin = 2;
    	            } else if (statusPrioritasMesin == 2) {
    	            	statusPrioritasMesin = 3;
    	            } else if (statusPrioritasMesin == 3) {
    	            	statusPrioritasMesin = 4;
    	            } else if (statusPrioritasMesin == 4) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 12) {
    	            	statusPrioritasMesin = 13;
    	            } else if (statusPrioritasMesin == 13) {
    	            	statusPrioritasMesin = 14;
    	            } else if (statusPrioritasMesin == 14) {
    	            	statusPrioritasMesin = 15;
    	            } else if(statusPrioritasMesin == 15) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 22) {
    	        		statusPrioritasMesin = 23;
    	            } else if (statusPrioritasMesin == 23) {
    	            	statusPrioritasMesin = 24;
    	            } else if (statusPrioritasMesin == 24) {
    	            	statusPrioritasMesin = 25;
    	            } else if (statusPrioritasMesin == 25) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 32) {
    	            	statusPrioritasMesin = 33;
    	            } else if (statusPrioritasMesin == 33) {
    	            	statusPrioritasMesin = 34;
    	            } else if (statusPrioritasMesin == 34) {
    	            	statusPrioritasMesin = 35;
    	            } else if(statusPrioritasMesin == 35) {
    	            	statusPrioritasMesin = 2;
    	            }  else if(statusPrioritasMesin == 99) {
    	            	break;
    	            } 
	            }	
            }
	    	machineCuringList = machineCuringListTemp;
            dtMo.setMoMonth0(order);	
		}
		
		for(DetailMo dtMo : detailMarketingOrderListBOMFrontRear) {
		////System.out.println("test 11");
			if(!checkAllActiveMachine()) {
				break;
			}
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 12");
            order = dtMo.getTotalAR();
           ////System.out.println("test 13");
            minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	    	while (order.intValue() > minProduction.intValue()) {
            ////System.out.println("test 14");
            	List<CTCuring> ctCurList = getMachine(statusPrioritasMesin, dtMo.getItemCuring());
	            if(ctCurList.size() > 0) {
	            ////System.out.println("test 15");
    	            tempShift = generateFromMidMonth(dtMo.getPartNumber(), month, year,  ctCurList);
	            }
	           ////System.out.println("test 16");
	            if(!tempShift) {
	            ////System.out.println("test 17");
	            	clearShift(dtMo.getPartNumber());
	            ////System.out.println("test 18");]
	            	if (statusPrioritasMesin == 1) {
    	        		statusPrioritasMesin = 2;
    	            } else if (statusPrioritasMesin == 2) {
    	            	statusPrioritasMesin = 3;
    	            } else if (statusPrioritasMesin == 3) {
    	            	statusPrioritasMesin = 4;
    	            } else if (statusPrioritasMesin == 4) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 12) {
    	            	statusPrioritasMesin = 13;
    	            } else if (statusPrioritasMesin == 13) {
    	            	statusPrioritasMesin = 14;
    	            } else if (statusPrioritasMesin == 14) {
    	            	statusPrioritasMesin = 15;
    	            } else if(statusPrioritasMesin == 15) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 22) {
    	        		statusPrioritasMesin = 23;
    	            } else if (statusPrioritasMesin == 23) {
    	            	statusPrioritasMesin = 24;
    	            } else if (statusPrioritasMesin == 24) {
    	            	statusPrioritasMesin = 25;
    	            } else if (statusPrioritasMesin == 25) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 32) {
    	            	statusPrioritasMesin = 33;
    	            } else if (statusPrioritasMesin == 33) {
    	            	statusPrioritasMesin = 34;
    	            } else if (statusPrioritasMesin == 34) {
    	            	statusPrioritasMesin = 35;
    	            } else if(statusPrioritasMesin == 35) {
    	            	statusPrioritasMesin = 2;
    	            }  else if(statusPrioritasMesin == 99) {
    	            	break;
    	            } 
	            }	
            }
	    	machineCuringList = machineCuringListTemp;
            dtMo.setMoMonth0(order);
			
		}
		
		for(DetailMo dtMo : detailMarketingOrderListDualFrontRear) {
		////System.out.println("test 11");
			if(!checkAllActiveMachine()) {
				break;
			}
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 12");
            order = dtMo.getTotalAR();
           ////System.out.println("test 13");
            minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	    	while (order.intValue() > minProduction.intValue()) {
            ////System.out.println("test 14");
            	List<CTCuring> ctCurList = getMachine(statusPrioritasMesin, dtMo.getItemCuring());
	            if(ctCurList.size() > 0) {
	            ////System.out.println("test 15");
    	            tempShift = generateFromMidMonth(dtMo.getPartNumber(), month, year,  ctCurList);
	            }
	           ////System.out.println("test 16");
	            if(!tempShift) {
	            ////System.out.println("test 17");
	            	clearShift(dtMo.getPartNumber());
	            ////System.out.println("test 18");
	            	if (statusPrioritasMesin == 1) {
    	        		statusPrioritasMesin = 2;
    	            } else if (statusPrioritasMesin == 2) {
    	            	statusPrioritasMesin = 3;
    	            } else if (statusPrioritasMesin == 3) {
    	            	statusPrioritasMesin = 4;
    	            } else if (statusPrioritasMesin == 4) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 12) {
    	            	statusPrioritasMesin = 13;
    	            } else if (statusPrioritasMesin == 13) {
    	            	statusPrioritasMesin = 14;
    	            } else if (statusPrioritasMesin == 14) {
    	            	statusPrioritasMesin = 15;
    	            } else if(statusPrioritasMesin == 15) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 22) {
    	        		statusPrioritasMesin = 23;
    	            } else if (statusPrioritasMesin == 23) {
    	            	statusPrioritasMesin = 24;
    	            } else if (statusPrioritasMesin == 24) {
    	            	statusPrioritasMesin = 25;
    	            } else if (statusPrioritasMesin == 25) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 32) {
    	            	statusPrioritasMesin = 33;
    	            } else if (statusPrioritasMesin == 33) {
    	            	statusPrioritasMesin = 34;
    	            } else if (statusPrioritasMesin == 34) {
    	            	statusPrioritasMesin = 35;
    	            } else if(statusPrioritasMesin == 35) {
    	            	statusPrioritasMesin = 2;
    	            }  else if(statusPrioritasMesin == 99) {
    	            	break;
    	            } 
	            }	
            }
	    	machineCuringList = machineCuringListTemp;
            dtMo.setMoMonth0(order);
			
		}
		
		for(DetailMo dtMo : detailMarketingOrderListAB) {
		////System.out.println("test 2");
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 3");
		    order = dtMo.getTotalAR();
		   ////System.out.println("test 4");
    	    oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), dtMo.getItemCuring());
    	   ////System.out.println("test 5");
    	    if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) { // if 1
    			minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
    	    	while (order.intValue() > minProduction.intValue()) {
        	    ////System.out.println("test 6");
        	    	if(!checkAllActiveMachine()) {
        	    		break;
        	    	}
        	    ////System.out.println("test 7");

    	            tempShift = generateFromOldShift(month, year);
    	            
        	       ////System.out.println("test 8");
        	        if(!tempShift){
        	        	break;
        	        }
        	    }
        	   ////System.out.println("test 9");
        	    dtMo.setMoMonth0(order);
	        }
    	    machineCuringList = machineCuringListTemp;
		}
		
		for(DetailMo dtMo : detailMarketingOrderListBOM) {
			////System.out.println("test 2");
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 3");
		    order = dtMo.getTotalAR();
		   ////System.out.println("test 4");
    	    oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), dtMo.getItemCuring());
    	   ////System.out.println("test 5");
    	    if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) { // if 1
    	    	minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
    	    	while (order.intValue() > minProduction.intValue()) {
        	    ////System.out.println("test 6");
        	    	if(!checkAllActiveMachine()) {
        	    		break;
        	    	}
        	    ////System.out.println("test 7");

    	            tempShift = generateFromOldShift(month, year);
    	            
        	       ////System.out.println("test 8");
        	        if(!tempShift){
        	        	break;
        	        }
        	    }
        	   ////System.out.println("test 9");
        	    dtMo.setMoMonth0(order);
	        }
    	    machineCuringList = machineCuringListTemp;
		}
		
		for(DetailMo dtMo : detailMarketingOrderListDual) {
		////System.out.println("test 2");
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 3");
		    order = dtMo.getTotalAR();
		   ////System.out.println("test 4");
    	    oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), dtMo.getItemCuring());
    	   ////System.out.println("test 5");
    	    if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) { // if 1
    	    	minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
    	    	while (order.intValue() > minProduction.intValue()) {
        	    ////System.out.println("test 6");
        	    	if(!checkAllActiveMachine()) {
        	    		break;
        	    	}
        	    	////System.out.println("test 7");

    	            tempShift = generateFromOldShift(month, year);
    	            
        	       ////System.out.println("test 8");
        	        if(!tempShift){
        	        	break;
        	        }
        	    }
        	   ////System.out.println("test 9");
        	    dtMo.setMoMonth0(order);
	        }
    	    machineCuringList = machineCuringListTemp;
		}
		
		for(DetailMo dtMo : detailMarketingOrderListAB) {
		////System.out.println("test 11");
			if(!checkAllActiveMachine()) {
				break;
			}
			int statusPrioritasMesin = 0;
			tempShift = false;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 12");
            order = dtMo.getTotalAR();
           ////System.out.println("test 13");
            minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	    	while (order.intValue() > minProduction.intValue()) {
            ////System.out.println("test 14");
            	List<CTCuring> ctCurList = getMachine(statusPrioritasMesin, dtMo.getItemCuring());
	            if(ctCurList.size() > 0) {
	            ////System.out.println("test 15");
    	            tempShift = generateFromMidMonth(dtMo.getPartNumber(), month, year,  ctCurList);
	            }
	           ////System.out.println("test 16");
	            if(!tempShift) {
	            ////System.out.println("test 17");
	            	clearShift(dtMo.getPartNumber());
	            ////System.out.println("test 18");
	            	if (statusPrioritasMesin == 1) {
    	        		statusPrioritasMesin = 2;
    	            } else if (statusPrioritasMesin == 2) {
    	            	statusPrioritasMesin = 3;
    	            } else if (statusPrioritasMesin == 3) {
    	            	statusPrioritasMesin = 4;
    	            } else if (statusPrioritasMesin == 4) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 12) {
    	            	statusPrioritasMesin = 13;
    	            } else if (statusPrioritasMesin == 13) {
    	            	statusPrioritasMesin = 14;
    	            } else if (statusPrioritasMesin == 14) {
    	            	statusPrioritasMesin = 15;
    	            } else if(statusPrioritasMesin == 15) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 22) {
    	        		statusPrioritasMesin = 23;
    	            } else if (statusPrioritasMesin == 23) {
    	            	statusPrioritasMesin = 24;
    	            } else if (statusPrioritasMesin == 24) {
    	            	statusPrioritasMesin = 25;
    	            } else if (statusPrioritasMesin == 25) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 32) {
    	            	statusPrioritasMesin = 33;
    	            } else if (statusPrioritasMesin == 33) {
    	            	statusPrioritasMesin = 34;
    	            } else if (statusPrioritasMesin == 34) {
    	            	statusPrioritasMesin = 35;
    	            } else if(statusPrioritasMesin == 35) {
    	            	statusPrioritasMesin = 2;
    	            }  else if(statusPrioritasMesin == 99) {
    	            	break;
    	            } 
	            }	
            }
	    	machineCuringList = machineCuringListTemp;
            dtMo.setMoMonth0(order);	
		}
		
		for(DetailMo dtMo : detailMarketingOrderListBOM) {
		////System.out.println("test 11");
			if(!checkAllActiveMachine()) {
				break;
			}
			tempShift = false;
			int statusPrioritasMesin = 0;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 12");
            order = dtMo.getTotalAR();
           ////System.out.println("test 13");
            minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	    	while (order.intValue() > minProduction.intValue()) {
            ////System.out.println("test 14");
            	List<CTCuring> ctCurList = getMachine(statusPrioritasMesin, dtMo.getItemCuring());
	            if(ctCurList.size() > 0) {
	            ////System.out.println("test 15");
    	            tempShift = generateFromMidMonth(dtMo.getPartNumber(), month, year,  ctCurList);
	            }
	           ////System.out.println("test 16");
	            if(!tempShift) {
	            ////System.out.println("test 17");
	            	clearShift(dtMo.getPartNumber());
	            ////System.out.println("test 18");
	            	if (statusPrioritasMesin == 1) {
    	        		statusPrioritasMesin = 2;
    	            } else if (statusPrioritasMesin == 2) {
    	            	statusPrioritasMesin = 3;
    	            } else if (statusPrioritasMesin == 3) {
    	            	statusPrioritasMesin = 4;
    	            } else if (statusPrioritasMesin == 4) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 12) {
    	            	statusPrioritasMesin = 13;
    	            } else if (statusPrioritasMesin == 13) {
    	            	statusPrioritasMesin = 14;
    	            } else if (statusPrioritasMesin == 14) {
    	            	statusPrioritasMesin = 15;
    	            } else if(statusPrioritasMesin == 15) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 22) {
    	        		statusPrioritasMesin = 23;
    	            } else if (statusPrioritasMesin == 23) {
    	            	statusPrioritasMesin = 24;
    	            } else if (statusPrioritasMesin == 24) {
    	            	statusPrioritasMesin = 25;
    	            } else if (statusPrioritasMesin == 25) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 32) {
    	            	statusPrioritasMesin = 33;
    	            } else if (statusPrioritasMesin == 33) {
    	            	statusPrioritasMesin = 34;
    	            } else if (statusPrioritasMesin == 34) {
    	            	statusPrioritasMesin = 35;
    	            } else if(statusPrioritasMesin == 35) {
    	            	statusPrioritasMesin = 2;
    	            }  else if(statusPrioritasMesin == 99) {
    	            	break;
    	            } 
	            }	
            }
	    	machineCuringList = machineCuringListTemp;
            dtMo.setMoMonth0(order);
			
		}
		
		for(DetailMo dtMo : detailMarketingOrderListDual) {
		////System.out.println("test 11");
			if(!checkAllActiveMachine()) {
				break;
			}
			tempShift = false;
			int statusPrioritasMesin = 0;
            statusPrioritasMesin = getStatusPrioritasMesin(dtMo, smallOrderLimit);
           ////System.out.println("test 12");
            order = dtMo.getTotalAR();
           ////System.out.println("test 13");
            minProduction = getMinimalProduction(dtMo.getTotalAR(), minA, minB, minC, minD);
	    	while (order.intValue() > minProduction.intValue()) {
            ////System.out.println("test 14");
            	List<CTCuring> ctCurList = getMachine(statusPrioritasMesin, dtMo.getItemCuring());
	            if(ctCurList.size() > 0) {
	            ////System.out.println("test 15");
    	            tempShift = generateFromMidMonth(dtMo.getPartNumber(), month, year,  ctCurList);
	            }
	           ////System.out.println("test 16");
	            if(!tempShift) {
	            ////System.out.println("test 17");
	            	clearShift(dtMo.getPartNumber());
	            ////System.out.println("test 18");
	            	if (statusPrioritasMesin == 1) {
    	        		statusPrioritasMesin = 2;
    	            } else if (statusPrioritasMesin == 2) {
    	            	statusPrioritasMesin = 3;
    	            } else if (statusPrioritasMesin == 3) {
    	            	statusPrioritasMesin = 4;
    	            } else if (statusPrioritasMesin == 4) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 12) {
    	            	statusPrioritasMesin = 13;
    	            } else if (statusPrioritasMesin == 13) {
    	            	statusPrioritasMesin = 14;
    	            } else if (statusPrioritasMesin == 14) {
    	            	statusPrioritasMesin = 15;
    	            } else if(statusPrioritasMesin == 15) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 22) {
    	        		statusPrioritasMesin = 23;
    	            } else if (statusPrioritasMesin == 23) {
    	            	statusPrioritasMesin = 24;
    	            } else if (statusPrioritasMesin == 24) {
    	            	statusPrioritasMesin = 25;
    	            } else if (statusPrioritasMesin == 25) {
    	            	statusPrioritasMesin = 99;
    	            } else if (statusPrioritasMesin == 32) {
    	            	statusPrioritasMesin = 33;
    	            } else if (statusPrioritasMesin == 33) {
    	            	statusPrioritasMesin = 34;
    	            } else if (statusPrioritasMesin == 34) {
    	            	statusPrioritasMesin = 35;
    	            } else if(statusPrioritasMesin == 35) {
    	            	statusPrioritasMesin = 2;
    	            }  else if(statusPrioritasMesin == 99) {
    	            	break;
    	            } 
	            }	
            }
	    	machineCuringList = machineCuringListTemp;
            dtMo.setMoMonth0(order);
			
		}
		
		for(ChangeMould obj : changeMouldList) {
			System.out.println(obj.getPartNum() + " " + obj.getChangeDate() + " " + obj.getWct() + " " + obj.getShift());
		}
		System.out.println("ini ab doang");
		for(DetailMo dtm : detailMarketingOrderListAB) {
			if(dtm.getTotalAR().intValue() > 0) {
				System.out.println(dtm.getPartNumber() + " " + dtm.getTotalAR());
			}
		}
		
		System.out.println("ini bom doang");
		for(DetailMo dtm : detailMarketingOrderListBOM) {
			if(dtm.getTotalAR().intValue() > 0) {
				System.out.println(dtm.getPartNumber() + " " + dtm.getTotalAR());
			}
		}
		
		System.out.println("ini dual ");
		for(DetailMo dtm : detailMarketingOrderListDual) {
			if(dtm.getTotalAR().intValue() > 0) {
				System.out.println(dtm.getPartNumber() + " " + dtm.getTotalAR());
			}
		}
		
		return newShiftMonthlyPlan;
	}
	
	public BigDecimal getMinimalProduction(BigDecimal mo, BigDecimal minA, BigDecimal minB, BigDecimal minC, BigDecimal minD) {
	    BigDecimal percentage;

	    if (mo.compareTo(BigDecimal.valueOf(2001)) < 0) {
	        percentage = minA;
	    } else if (mo.compareTo(BigDecimal.valueOf(10001)) < 0) {
	        percentage = minB;
	    } else if (mo.compareTo(BigDecimal.valueOf(100001)) < 0) {
	        percentage = minC;
	    } else {
	        percentage = minD;
	    }

	    if (percentage.compareTo(BigDecimal.ZERO) != 0) {
	        return mo.multiply(percentage).divide(BigDecimal.valueOf(100), RoundingMode.HALF_UP);
	    }

	    return BigDecimal.ZERO;
	}
	
	public BigDecimal getMaxProduction(BigDecimal mo, BigDecimal maxA, BigDecimal maxB, BigDecimal maxC, BigDecimal maxD) {
	    BigDecimal percentage;

	    if (mo.compareTo(BigDecimal.valueOf(2001)) < 0) {
	        percentage = maxA;
	    } else if (mo.compareTo(BigDecimal.valueOf(10001)) < 0) {
	        percentage = maxB;
	    } else if (mo.compareTo(BigDecimal.valueOf(100001)) < 0) {
	        percentage = maxC;
	    } else {
	        percentage = maxD;
	    }
	    
	    BigDecimal extra = mo.add(mo.multiply(percentage).divide(BigDecimal.valueOf(100), RoundingMode.HALF_UP));
	    if (percentage.compareTo(BigDecimal.ZERO) != 0) {
	        return extra;
	    }

	    return BigDecimal.ZERO;
	}
	
	public boolean checkMo() {
		for(DetailMo dtMo : detailMarketingOrderList) {
			if (dtMo.getTotalAR().compareTo(dtMo.getProductionLimit()) > 0) {
				return true;
			}
		}
		return false;
	}
	
	public boolean validateChangeMould(Date dateChange, int shift) {
		int i = 0;
		for(ChangeMould obj : endProductList) {
			if(obj.getChangeDate().equals(dateChange) && obj.getShift() == shift) {
				i++;
			}
		}
		if(i < maxChangeMould) {
			return true;
		}
		return false;
	}
	//4338
	
	public boolean generateFromManualMapping(BigDecimal partNum, int month, int year, String itemCuring, String wct) {
		int cav = 0;
	////System.out.println("ini minimal bikin nya " + minProduction);
		for (MachineCuring machineCuring : machineCuringListTemp) {
			if(machineCuring.getWORK_CENTER_TEXT().equals(wct)) {
				cav = machineCuring.getCAVITY().intValue();
				List<Map<String, Object>> list = new ArrayList<>();
				if(checkOldShift(itemCuring)) {
				////System.out.println("masuk old shift");
					list = dWorkDayHourSpecificRepo.getCuringCapacity(itemCuring , wct, cav, month, year);
				}else {
				////System.out.println("masuk old shift2");
					list = dWorkDayHourSpecificRepo.getCuringCapacityChangeMouldFirstDate(itemCuring , wct, cav, month, year);
					addChangeMould(workDayList.get(0).getDATE_WD(), partNum , 1, machineCuring.getWORK_CENTER_TEXT() +  " dari manual mapping");
				}
				if (list != null && !list.isEmpty()) {
					for(Map<String, Object> capacityData : list) {
		    			////System.out.println("masuk 8 " + new BigDecimal(capacityData.get("SHIFT1_CAPACITY").toString()) + " tanggal " + capacityData.get("DATE_WD").toString());
						if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
							BigDecimal tempShift = BigDecimal.ZERO;
							ShiftMonthlyPlan shift = new ShiftMonthlyPlan();
		                    shift.setKAPA_SHIFT_1(BigDecimal.ZERO);
		                    shift.setKAPA_SHIFT_2(BigDecimal.ZERO);
		                    shift.setKAPA_SHIFT_3(BigDecimal.ZERO);
		                    shift.setDATE(parseDate(capacityData.get("DATE_WD").toString()));
		                    shift.setPART_NUMBER(partNum);
		                    shift.setCAVITY(new BigDecimal(cav));
		                    shift.setWORK_CENTER_TEXT(wct);
		                    shift.setCAVITY_USAGE(new BigDecimal(cav));
		                    shift.setSTATUS(BigDecimal.ONE);
		                    shift.setCAVITY_EXIST(BigDecimal.ZERO);
		                    shift.setITEM_CURING(itemCuring);
		                    if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
		                    	shift.setWH_SHIFT_1(new BigDecimal(capacityData.get("SHIFT1_HOUR").toString()));
		                        shift.setKAPA_SHIFT_1(new BigDecimal(capacityData.get("SHIFT1_CAPACITY").toString()));	
		                        order = order.subtract(shift.getKAPA_SHIFT_1());
		                        tempShift = BigDecimal.ONE;
		                    }
		                    if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
		                    	shift.setWH_SHIFT_2(new BigDecimal(capacityData.get("SHIFT2_HOUR").toString()));
		                        shift.setKAPA_SHIFT_2(new BigDecimal(capacityData.get("SHIFT2_CAPACITY").toString()));	
		                        order = order.subtract(shift.getKAPA_SHIFT_2());
		                        tempShift = BigDecimal.valueOf(2);
		                    }
		                    if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
		                    	shift.setWH_SHIFT_3(new BigDecimal(capacityData.get("SHIFT3_HOUR").toString()));
		                        shift.setKAPA_SHIFT_3(new BigDecimal(capacityData.get("SHIFT3_CAPACITY").toString()));	
		                        order = order.subtract(shift.getKAPA_SHIFT_3());
		                        tempShift = BigDecimal.valueOf(3);
		                    }
		                    BigDecimal totalKapasitasShift = shift.getKAPA_SHIFT_1()
		                            .add(shift.getKAPA_SHIFT_2())
		                            .add(shift.getKAPA_SHIFT_3());
		                    shift.setTOTAL_KAPA(totalKapasitasShift);
		                    newShiftMonthlyPlan.add(shift);
		        			////System.out.println("masuk 9");
		        			if(order.intValue() < minProduction.intValue() && validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
		        				machineCuring.setSTATUS_USAGE(tempShift);
		        				machineCuring.setLAST_UPDATE_DATE(parseDate(capacityData.get("DATE_WD").toString()));
		        				addEndMould(parseDate(capacityData.get("DATE_WD").toString()), partNum , tempShift.intValue(), machineCuring.getWORK_CENTER_TEXT(), 0);
		                    	return true;
		                    }
						}
		    			////System.out.println("masuk 10");
					}
					machineCuring.setSTATUS(BigDecimal.ZERO);
					return true;
		        }
				break;
			}
		}
		return false;
	}
	
	public boolean checkOldShift(String itemCuring) {
		oldShiftPlan = shiftMonthlyRepo.findYesterdaysShiftPlan(formatDateToString(workDayList.get(0).getDATE_WD()), itemCuring);
		if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) {
			return true;
		}
		return false;
	}
	
	public boolean generateFromMidMonth(BigDecimal partNum, int month, int year, List<CTCuring> ctCurList) {
	////System.out.println("Masuk ke mid month0");
		for (CTCuring ctCur : ctCurList) {
		////System.out.println("Masuk ke mid month1");
            for (MachineCuring machineCuring : machineCuringListTemp) {
                if (machineCuring.getWORK_CENTER_TEXT().equals(ctCur.getOPERATION_SHORT_TEXT()) && machineCuring.getSTATUS().equals(BigDecimal.ONE)) {
                ////System.out.println("mesin " + machineCuring.getLAST_UPDATE_DATE() + " " + machineCuring.getSTATUS().intValue());
                	//if(validateChangeMould(machineCuring.getLAST_UPDATE_DATE(), machineCuring.getSTATUS_USAGE().intValue())) {
                		List<Map<String, Object>> list = new ArrayList<>();
                		if(machineCuring.getLAST_UPDATE_DATE() == null) {
                			list = dWorkDayHourSpecificRepo.getCuringCapacityChangeMouldFirstDate(ctCur.getWIP() , ctCur.getOPERATION_SHORT_TEXT(), machineCuring.getCAVITY().intValue(), month, year);
                		}else {
                			list = dWorkDayHourSpecificRepo.getCuringCapacityMidMonth(ctCur.getWIP() , ctCur.getOPERATION_SHORT_TEXT(), machineCuring.getCAVITY().intValue(), formatDateToString(machineCuring.getLAST_UPDATE_DATE()), machineCuring.getSTATUS_USAGE().intValue());	
                		}
                	////System.out.println("ukuran list " + list.size() + " " + ctCur.getWIP() + " " + ctCur.getOPERATION_SHORT_TEXT());
    					if (list != null && !list.isEmpty()) {
    					////System.out.println("masuk list tidak sama dengan null " + ctCur.getWIP());
    						addChangeMould(machineCuring.getLAST_UPDATE_DATE(), partNum , machineCuring.getSTATUS_USAGE().intValue(), machineCuring.getWORK_CENTER_TEXT() + " dari mid month");
    						for(Map<String, Object> capacityData : list) {
    	            			////System.out.println("masuk 8 " + new BigDecimal(capacityData.get("SHIFT1_CAPACITY").toString()) + " tanggal " + capacityData.get("DATE_WD").toString());
    							if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
    								BigDecimal tempShift = BigDecimal.ZERO;
    								ShiftMonthlyPlan shift = new ShiftMonthlyPlan();
    		                        shift.setKAPA_SHIFT_1(BigDecimal.ZERO);
    		                        shift.setKAPA_SHIFT_2(BigDecimal.ZERO);
    		                        shift.setKAPA_SHIFT_3(BigDecimal.ZERO);
    		                        shift.setDATE(parseDate(capacityData.get("DATE_WD").toString()));
    		                        shift.setPART_NUMBER(partNum);
    		                        shift.setCAVITY(machineCuring.getCAVITY());
    		                        shift.setWORK_CENTER_TEXT(machineCuring.getWORK_CENTER_TEXT());
    		                        shift.setCAVITY_USAGE(machineCuring.getCAVITY());
    		                        shift.setCAVITY_EXIST(BigDecimal.ZERO);
    		                        shift.setITEM_CURING(ctCur.getWIP());
    		                        if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
    		                        	shift.setWH_SHIFT_1(new BigDecimal(capacityData.get("SHIFT1_HOUR").toString()));
    			                        shift.setKAPA_SHIFT_1(new BigDecimal(capacityData.get("SHIFT1_CAPACITY").toString()));	
    			                        order = order.subtract(shift.getKAPA_SHIFT_1());
    			                        tempShift = BigDecimal.ONE;
    		                        }
    		                        if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
    		                        	shift.setWH_SHIFT_2(new BigDecimal(capacityData.get("SHIFT2_HOUR").toString()));
    			                        shift.setKAPA_SHIFT_2(new BigDecimal(capacityData.get("SHIFT2_CAPACITY").toString()));	
    			                        order = order.subtract(shift.getKAPA_SHIFT_2());
    			                        tempShift = BigDecimal.valueOf(2);
    		                        }
    		                        if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
    		                        	shift.setWH_SHIFT_3(new BigDecimal(capacityData.get("SHIFT3_HOUR").toString()));
    			                        shift.setKAPA_SHIFT_3(new BigDecimal(capacityData.get("SHIFT3_CAPACITY").toString()));	
    			                        order = order.subtract(shift.getKAPA_SHIFT_3());
    			                        tempShift = BigDecimal.valueOf(3);
    		                        }
    		                        BigDecimal totalKapasitasShift = shift.getKAPA_SHIFT_1()
    		                                .add(shift.getKAPA_SHIFT_2())
    		                                .add(shift.getKAPA_SHIFT_3());
    		                        shift.setTOTAL_KAPA(totalKapasitasShift);
    		                        newShiftMonthlyPlan.add(shift);
    		            			////System.out.println("masuk 9");
    		            			if(order.intValue() < minProduction.intValue() && validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
    		            				machineCuring.setSTATUS_USAGE(tempShift);
    		            				machineCuring.setLAST_UPDATE_DATE(parseDate(capacityData.get("DATE_WD").toString()));
    		            				addEndMould(parseDate(capacityData.get("DATE_WD").toString()), partNum , tempShift.intValue(), machineCuring.getWORK_CENTER_TEXT(), 1);
    		                        	return true;
    		                        }
    							}
    	            			////System.out.println("masuk 10");
    						}
    					////System.out.println("habis " + machineCuring.getWORK_CENTER_TEXT());
							machineCuring.setSTATUS(BigDecimal.ZERO);
    						return true;
                        //}
                	}else {
                		machineCuring.setSTATUS(BigDecimal.ZERO);
                	}
                	
                }
            }
        }
		
		return false;
	}
	
	public boolean generateFromOldShift(int month, int year) {
		if (oldShiftPlan != null && !oldShiftPlan.isEmpty()) {
			for (ShiftMonthlyPlan shiftPlan : oldShiftPlan) {
                for (MachineCuring machineCuring : machineCuringListTemp) {
                    if (shiftPlan.getWORK_CENTER_TEXT().equals(machineCuring.getWORK_CENTER_TEXT()) && machineCuring.getSTATUS().equals(BigDecimal.ONE) && (machineCuring.getLAST_UPDATE_DATE() == null)) {
                    	List<Map<String, Object>> list = dWorkDayHourSpecificRepo.getCuringCapacity(shiftPlan.getITEM_CURING() , shiftPlan.getWORK_CENTER_TEXT(), shiftPlan.getCAVITY().intValue(), month, year);
						if (list != null && !list.isEmpty()) {
							for(Map<String, Object> capacityData : list) {
		            			////System.out.println("masuk 8 " + new BigDecimal(capacityData.get("SHIFT1_CAPACITY").toString()) + " tanggal " + capacityData.get("DATE_WD").toString());
								if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
									BigDecimal tempShift = BigDecimal.ZERO;
									ShiftMonthlyPlan shift = new ShiftMonthlyPlan();
			                        shift.setKAPA_SHIFT_1(BigDecimal.ZERO);
			                        shift.setKAPA_SHIFT_2(BigDecimal.ZERO);
			                        shift.setKAPA_SHIFT_3(BigDecimal.ZERO);
			                        shift.setDATE(parseDate(capacityData.get("DATE_WD").toString()));
			                        shift.setPART_NUMBER(shiftPlan.getPART_NUMBER());
			                        shift.setCAVITY(shiftPlan.getCAVITY());
			                        shift.setWORK_CENTER_TEXT(shiftPlan.getWORK_CENTER_TEXT());
			                        shift.setCAVITY_USAGE(shiftPlan.getCAVITY());
			                        shift.setSTATUS(BigDecimal.ONE);
			                        shift.setCAVITY_EXIST(BigDecimal.ZERO);
			                        shift.setITEM_CURING(shiftPlan.getITEM_CURING());
			                        if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
			                        	shift.setWH_SHIFT_1(new BigDecimal(capacityData.get("SHIFT1_HOUR").toString()));
				                        shift.setKAPA_SHIFT_1(new BigDecimal(capacityData.get("SHIFT1_CAPACITY").toString()));	
				                        order = order.subtract(shift.getKAPA_SHIFT_1());
				                        tempShift = BigDecimal.ONE;
			                        }
			                        if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
			                        	shift.setWH_SHIFT_2(new BigDecimal(capacityData.get("SHIFT2_HOUR").toString()));
				                        shift.setKAPA_SHIFT_2(new BigDecimal(capacityData.get("SHIFT2_CAPACITY").toString()));	
				                        order = order.subtract(shift.getKAPA_SHIFT_2());
				                        tempShift = BigDecimal.valueOf(2);
			                        }
			                        if(order.intValue() > minProduction.intValue() || !validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
			                        	shift.setWH_SHIFT_3(new BigDecimal(capacityData.get("SHIFT3_HOUR").toString()));
				                        shift.setKAPA_SHIFT_3(new BigDecimal(capacityData.get("SHIFT3_CAPACITY").toString()));	
				                        order = order.subtract(shift.getKAPA_SHIFT_3());
				                        tempShift = BigDecimal.valueOf(3);
			                        }
			                        BigDecimal totalKapasitasShift = shift.getKAPA_SHIFT_1()
			                                .add(shift.getKAPA_SHIFT_2())
			                                .add(shift.getKAPA_SHIFT_3());
			                        shift.setTOTAL_KAPA(totalKapasitasShift);
			                        newShiftMonthlyPlan.add(shift);
			            			////System.out.println("masuk 9");
			            			if(order.intValue() < minProduction.intValue() && validateChangeMould(parseDate(capacityData.get("DATE_WD").toString()), 1)) {
			            				machineCuring.setSTATUS_USAGE(tempShift);
			            				machineCuring.setLAST_UPDATE_DATE(parseDate(capacityData.get("DATE_WD").toString()));
			            				addEndMould(parseDate(capacityData.get("DATE_WD").toString()), shiftPlan.getPART_NUMBER() , tempShift.intValue(), machineCuring.getWORK_CENTER_TEXT(), 0);
			                        	return true;
			                        }
								}
		            			////System.out.println("masuk 10");
							}
							machineCuring.setSTATUS(BigDecimal.ZERO);
							return true;
	                    }
                    }
                }
            }
		}
		
		return false;
	}
  	
	public void addChangeMould(Date cahangeDate, BigDecimal partNum, int shift, String wct) {
		ChangeMould obj = new ChangeMould();
		obj.setChangeDate(cahangeDate);
		obj.setPartNum(partNum);
		obj.setShift(shift);
		obj.setWct(wct);
		changeMouldList.add(obj);
	}
	
	public void addEndMould(Date cahangeDate, BigDecimal partNum, int shift, String wct, int status) {
		ChangeMould obj = new ChangeMould();
		obj.setChangeDate(cahangeDate);
		obj.setPartNum(partNum);
		obj.setShift(shift);
		obj.setWct(wct);
		obj.setStatus(status);
		endProductList.add(obj);
	}
	
	public void clearDetailMarketingBatch1() {
	    Iterator<Map<String, Object>> iterator = detailMarkOrderList.iterator();

	    while (iterator.hasNext()) {
	        Map<String, Object> dtMo = iterator.next();
	        BigDecimal partNumber = new BigDecimal(dtMo.get("PART_NUMBER").toString());

	        for (TempOrder tempO : tempOrderList) {
	            if (partNumber.equals(tempO.getPratNum()) && tempO.getMarketingOrder().intValue() < 0) {
	                iterator.remove(); // Menghapus elemen dengan aman
	                break; // Hentikan iterasi tempOrderList setelah ditemukan
	            }
	        }
	    }
	}
	
	public BigDecimal getMarkOrderExist(BigDecimal partNum) {
		for (TempOrder tempO : tempOrderList) {
			if(tempO.getPratNum().equals(partNum)){
				return tempO.getMarketingOrder();
			}
		}
		return BigDecimal.ZERO;
	}
	
	public boolean checkActiveMachine(String wct) {
		for(MachineCuring mc : machineCuringUsedList) {
			if(mc.getWORK_CENTER_TEXT().equals(wct) && mc.getSTATUS_USAGE().equals(BigDecimal.ZERO)) {
				return false;
			}
		}
		return true;
	}
	
	public void getShiftByPartnumber(BigDecimal partnum){
		newShiftListPartNum.clear();
		for(ShiftMonthlyPlan sf : newShiftList) {
			if(sf.getPART_NUMBER().equals(partnum)) {
				newShiftListPartNum.add(sf);
			}
		}
	}
	
	public boolean checkAllActiveMachine() {
		for(MachineCuring mc : machineCuringList) {
			if(mc.getSTATUS().equals(BigDecimal.ONE)) {
				return true;
			}
		}
		
		return false;
	}
	
	public void getDataHeader(int month, int year) {
		machineCuringList = machineCuringRepo.findMachineCuringActive();
		machineCuringListTemp = machineCuringRepo.findMachineCuringActive();
		machineProductList = machineProductRepo.findAll();
		System.out.println("ukuran list mesin " + machineCuringList.size() + " " + machineCuringListTemp.size());
		 
    	////System.out.println(machineCuringList.size());
    	workDayList = workDayRepo.findByMonthYear(month, year); //Flowchart 3
		
    	smallOrderLimit = new BigDecimal(settingRepo.findSmallOrderLimit().getSETTING_VALUE());
    	
    	List<MarketingOrder> marketingOrderList = marketingOrderRepo.findByMonthYear(month, year); //flowchart 4

    	 //flowchart 5 6 7 8
    	////System.out.println("Done flow 5");
    	////System.out.println(marketingOrderList.get(0).getMoId() + " " + marketingOrderList.get(1).getMoId());
    	
    ////System.out.println("check");
    	List<Map<String, Object>> detailMarkOrderListAB = new ArrayList<>();
    ////System.out.println("check1");
    	List<Map<String, Object>> detailMarkOrderListBOM = new ArrayList<>();
    ////System.out.println("check2");
    	List<Map<String, Object>> detailMarkOrderListDual = new ArrayList<>();
    ////System.out.println("check3");
    	List<Map<String, Object>> detailMarkOrderListABFrontRear = new ArrayList<>();
    ////System.out.println("check1");
    	List<Map<String, Object>> detailMarkOrderListBOMFrontRear = new ArrayList<>();
    ////System.out.println("check2");
    	List<Map<String, Object>> detailMarkOrderListDualFrontRear = new ArrayList<>();

    	detailMarkOrderListAB = detailMarketingOrderRepo.findByMoIdSortProductTypeAbNotFrontRear(marketingOrderList.get(0).getMoId(), marketingOrderList.get(1).getMoId());
    ////System.out.println("check4");
    	detailMarkOrderListBOM = detailMarketingOrderRepo.findByMoIdSortProductTypeBomNotFrontRear(marketingOrderList.get(0).getMoId(), marketingOrderList.get(1).getMoId());
    ////System.out.println("check5");
    	detailMarkOrderListDual = detailMarketingOrderRepo.findByMoIdSortProductTypeBomAbNotFrontRear(marketingOrderList.get(0).getMoId(), marketingOrderList.get(1).getMoId());
    	
    	detailMarkOrderListABFrontRear = detailMarketingOrderRepo.findByMoIdSortProductTypeAbFrontRear(marketingOrderList.get(0).getMoId(), marketingOrderList.get(1).getMoId());
        ////System.out.println("check4");
    	detailMarkOrderListBOMFrontRear = detailMarketingOrderRepo.findByMoIdSortProductTypeBomAbFrontRear(marketingOrderList.get(0).getMoId(), marketingOrderList.get(1).getMoId());
    ////System.out.println("check5");
    	detailMarkOrderListDualFrontRear = detailMarketingOrderRepo.findByMoIdSortProductTypeBomFrontRear(marketingOrderList.get(0).getMoId(), marketingOrderList.get(1).getMoId());
    ////System.out.println("check6");
    System.out.println("ukuran  " + detailMarkOrderListAB.size() + " " + detailMarkOrderListBOM.size() + " " + detailMarkOrderListDual.size()+ " " + detailMarkOrderListABFrontRear.size() + " " + detailMarkOrderListBOMFrontRear.size() + " " + detailMarkOrderListDualFrontRear.size());
    	////System.out.println("Done flow 6");
        for (Map<String, Object> map : detailMarkOrderListAB) {
        	DetailMo obj = new DetailMo();
            obj.setPpd((BigDecimal) map.get("PPD"));
            obj.setLowerConstant((BigDecimal) map.get("LOWER_CONSTANT"));
            obj.setMoMonth0((BigDecimal) map.get("MO_MONTH_0"));
            obj.setProductCategory((String) map.get("PRODUCT_CATEGORY"));
            obj.setMachineType((String) map.get("MACHINE_TYPE"));
            obj.setItemExt((String) map.get("ITEM_EXT"));
            obj.setExtDescription((String) map.get("EXT_DESCRIPTION"));
            obj.setMinOrder((BigDecimal) map.get("MIN_ORDER"));
            obj.setSizeId((String) map.get("SIZE_ID"));
            obj.setDetailId((BigDecimal) map.get("DETAIL_ID"));
            obj.setRim((BigDecimal) map.get("RIM"));
            obj.setItemCuring((String) map.get("ITEM_CURING"));
            obj.setWibTube((String) map.get("WIB_TUBE"));
            obj.setPatternId((String) map.get("PATTERN_ID"));
            obj.setItemAssy((String) map.get("ITEM_ASSY"));
            obj.setCategory((String) map.get("CATEGORY"));
            obj.setCav((BigDecimal) map.get("CAV"));
            obj.setQtyPerRak((BigDecimal) map.get("QTY_PER_RAK"));
            obj.setDescription((String) map.get("DESCRIPTION"));
            obj.setQtyPerMould((BigDecimal) map.get("QTY_PER_MOULD"));
            obj.setMoId((String) map.get("MO_ID"));
            obj.setInitialStock((BigDecimal) map.get("INITIAL_STOCK"));
            obj.setCapacity((BigDecimal) map.get("CAPACITY"));
            obj.setProductType((String) map.get("PRODUCT_TYPE"));
            obj.setPartNumber((BigDecimal) map.get("PART_NUMBER"));
            obj.setUpperConstant((BigDecimal) map.get("UPPER_CONSTANT"));
            obj.setMaxCapMonth0((BigDecimal) map.get("MAX_CAP_MONTH_0"));
            detailMarketingOrderListAB.add(obj);
        }
        
        for (Map<String, Object> map : detailMarkOrderListBOM) {
        	DetailMo obj = new DetailMo();
            obj.setPpd((BigDecimal) map.get("PPD"));
            obj.setLowerConstant((BigDecimal) map.get("LOWER_CONSTANT"));
            obj.setMoMonth0((BigDecimal) map.get("MO_MONTH_0"));
            obj.setProductCategory((String) map.get("PRODUCT_CATEGORY"));
            obj.setMachineType((String) map.get("MACHINE_TYPE"));
            obj.setItemExt((String) map.get("ITEM_EXT"));
            obj.setExtDescription((String) map.get("EXT_DESCRIPTION"));
            obj.setMinOrder((BigDecimal) map.get("MIN_ORDER"));
            obj.setSizeId((String) map.get("SIZE_ID"));
            obj.setDetailId((BigDecimal) map.get("DETAIL_ID"));
            obj.setRim((BigDecimal) map.get("RIM"));
            obj.setItemCuring((String) map.get("ITEM_CURING"));
            obj.setWibTube((String) map.get("WIB_TUBE"));
            obj.setPatternId((String) map.get("PATTERN_ID"));
            obj.setItemAssy((String) map.get("ITEM_ASSY"));
            obj.setCategory((String) map.get("CATEGORY"));
            obj.setCav((BigDecimal) map.get("CAV"));
            obj.setQtyPerRak((BigDecimal) map.get("QTY_PER_RAK"));
            obj.setDescription((String) map.get("DESCRIPTION"));
            obj.setQtyPerMould((BigDecimal) map.get("QTY_PER_MOULD"));
            obj.setMoId((String) map.get("MO_ID"));
            obj.setInitialStock((BigDecimal) map.get("INITIAL_STOCK"));
            obj.setCapacity((BigDecimal) map.get("CAPACITY"));
            obj.setProductType((String) map.get("PRODUCT_TYPE"));
            obj.setPartNumber((BigDecimal) map.get("PART_NUMBER"));
            obj.setUpperConstant((BigDecimal) map.get("UPPER_CONSTANT"));
            obj.setMaxCapMonth0((BigDecimal) map.get("MAX_CAP_MONTH_0"));
            detailMarketingOrderListBOM.add(obj);
        }
        
        for (Map<String, Object> map : detailMarkOrderListDual) {
        	DetailMo obj = new DetailMo();
            obj.setPpd((BigDecimal) map.get("PPD"));
            obj.setLowerConstant((BigDecimal) map.get("LOWER_CONSTANT"));
            obj.setMoMonth0((BigDecimal) map.get("MO_MONTH_0"));
            obj.setProductCategory((String) map.get("PRODUCT_CATEGORY"));
            obj.setMachineType((String) map.get("MACHINE_TYPE"));
            obj.setItemExt((String) map.get("ITEM_EXT"));
            obj.setExtDescription((String) map.get("EXT_DESCRIPTION"));
            obj.setMinOrder((BigDecimal) map.get("MIN_ORDER"));
            obj.setSizeId((String) map.get("SIZE_ID"));
            obj.setDetailId((BigDecimal) map.get("DETAIL_ID"));
            obj.setRim((BigDecimal) map.get("RIM"));
            obj.setItemCuring((String) map.get("ITEM_CURING"));
            obj.setWibTube((String) map.get("WIB_TUBE"));
            obj.setPatternId((String) map.get("PATTERN_ID"));
            obj.setItemAssy((String) map.get("ITEM_ASSY"));
            obj.setCategory((String) map.get("CATEGORY"));
            obj.setCav((BigDecimal) map.get("CAV"));
            obj.setQtyPerRak((BigDecimal) map.get("QTY_PER_RAK"));
            obj.setDescription((String) map.get("DESCRIPTION"));
            obj.setQtyPerMould((BigDecimal) map.get("QTY_PER_MOULD"));
            obj.setMoId((String) map.get("MO_ID"));
            obj.setInitialStock((BigDecimal) map.get("INITIAL_STOCK"));
            obj.setCapacity((BigDecimal) map.get("CAPACITY"));
            obj.setProductType((String) map.get("PRODUCT_TYPE"));
            obj.setPartNumber((BigDecimal) map.get("PART_NUMBER"));
            obj.setUpperConstant((BigDecimal) map.get("UPPER_CONSTANT"));
            obj.setMaxCapMonth0((BigDecimal) map.get("MAX_CAP_MONTH_0"));
            detailMarketingOrderListDual.add(obj);
        }
        
        for (Map<String, Object> map : detailMarkOrderListABFrontRear) {
        	DetailMo obj = new DetailMo();
            obj.setPpd((BigDecimal) map.get("PPD"));
            obj.setLowerConstant((BigDecimal) map.get("LOWER_CONSTANT"));
            obj.setMoMonth0((BigDecimal) map.get("MO_MONTH_0"));
            obj.setProductCategory((String) map.get("PRODUCT_CATEGORY"));
            obj.setMachineType((String) map.get("MACHINE_TYPE"));
            obj.setItemExt((String) map.get("ITEM_EXT"));
            obj.setExtDescription((String) map.get("EXT_DESCRIPTION"));
            obj.setMinOrder((BigDecimal) map.get("MIN_ORDER"));
            obj.setSizeId((String) map.get("SIZE_ID"));
            obj.setDetailId((BigDecimal) map.get("DETAIL_ID"));
            obj.setRim((BigDecimal) map.get("RIM"));
            obj.setItemCuring((String) map.get("ITEM_CURING"));
            obj.setWibTube((String) map.get("WIB_TUBE"));
            obj.setPatternId((String) map.get("PATTERN_ID"));
            obj.setItemAssy((String) map.get("ITEM_ASSY"));
            obj.setCategory((String) map.get("CATEGORY"));
            obj.setCav((BigDecimal) map.get("CAV"));
            obj.setQtyPerRak((BigDecimal) map.get("QTY_PER_RAK"));
            obj.setDescription((String) map.get("DESCRIPTION"));
            obj.setQtyPerMould((BigDecimal) map.get("QTY_PER_MOULD"));
            obj.setMoId((String) map.get("MO_ID"));
            obj.setInitialStock((BigDecimal) map.get("INITIAL_STOCK"));
            obj.setCapacity((BigDecimal) map.get("CAPACITY"));
            obj.setProductType((String) map.get("PRODUCT_TYPE"));
            obj.setPartNumber((BigDecimal) map.get("PART_NUMBER"));
            obj.setUpperConstant((BigDecimal) map.get("UPPER_CONSTANT"));
            obj.setMaxCapMonth0((BigDecimal) map.get("MAX_CAP_MONTH_0"));
            detailMarketingOrderListABFrontRear.add(obj);
        }
        
        for (Map<String, Object> map : detailMarkOrderListBOMFrontRear) {
        	DetailMo obj = new DetailMo();
            obj.setPpd((BigDecimal) map.get("PPD"));
            obj.setLowerConstant((BigDecimal) map.get("LOWER_CONSTANT"));
            obj.setMoMonth0((BigDecimal) map.get("MO_MONTH_0"));
            obj.setProductCategory((String) map.get("PRODUCT_CATEGORY"));
            obj.setMachineType((String) map.get("MACHINE_TYPE"));
            obj.setItemExt((String) map.get("ITEM_EXT"));
            obj.setExtDescription((String) map.get("EXT_DESCRIPTION"));
            obj.setMinOrder((BigDecimal) map.get("MIN_ORDER"));
            obj.setSizeId((String) map.get("SIZE_ID"));
            obj.setDetailId((BigDecimal) map.get("DETAIL_ID"));
            obj.setRim((BigDecimal) map.get("RIM"));
            obj.setItemCuring((String) map.get("ITEM_CURING"));
            obj.setWibTube((String) map.get("WIB_TUBE"));
            obj.setPatternId((String) map.get("PATTERN_ID"));
            obj.setItemAssy((String) map.get("ITEM_ASSY"));
            obj.setCategory((String) map.get("CATEGORY"));
            obj.setCav((BigDecimal) map.get("CAV"));
            obj.setQtyPerRak((BigDecimal) map.get("QTY_PER_RAK"));
            obj.setDescription((String) map.get("DESCRIPTION"));
            obj.setQtyPerMould((BigDecimal) map.get("QTY_PER_MOULD"));
            obj.setMoId((String) map.get("MO_ID"));
            obj.setInitialStock((BigDecimal) map.get("INITIAL_STOCK"));
            obj.setCapacity((BigDecimal) map.get("CAPACITY"));
            obj.setProductType((String) map.get("PRODUCT_TYPE"));
            obj.setPartNumber((BigDecimal) map.get("PART_NUMBER"));
            obj.setUpperConstant((BigDecimal) map.get("UPPER_CONSTANT"));
            obj.setMaxCapMonth0((BigDecimal) map.get("MAX_CAP_MONTH_0"));
            detailMarketingOrderListBOMFrontRear.add(obj);
        }
        
        for (Map<String, Object> map : detailMarkOrderListDualFrontRear) {
        	DetailMo obj = new DetailMo();
            obj.setPpd((BigDecimal) map.get("PPD"));
            obj.setLowerConstant((BigDecimal) map.get("LOWER_CONSTANT"));
            obj.setMoMonth0((BigDecimal) map.get("MO_MONTH_0"));
            obj.setProductCategory((String) map.get("PRODUCT_CATEGORY"));
            obj.setMachineType((String) map.get("MACHINE_TYPE"));
            obj.setItemExt((String) map.get("ITEM_EXT"));
            obj.setExtDescription((String) map.get("EXT_DESCRIPTION"));
            obj.setMinOrder((BigDecimal) map.get("MIN_ORDER"));
            obj.setSizeId((String) map.get("SIZE_ID"));
            obj.setDetailId((BigDecimal) map.get("DETAIL_ID"));
            obj.setRim((BigDecimal) map.get("RIM"));
            obj.setItemCuring((String) map.get("ITEM_CURING"));
            obj.setWibTube((String) map.get("WIB_TUBE"));
            obj.setPatternId((String) map.get("PATTERN_ID"));
            obj.setItemAssy((String) map.get("ITEM_ASSY"));
            obj.setCategory((String) map.get("CATEGORY"));
            obj.setCav((BigDecimal) map.get("CAV"));
            obj.setQtyPerRak((BigDecimal) map.get("QTY_PER_RAK"));
            obj.setDescription((String) map.get("DESCRIPTION"));
            obj.setQtyPerMould((BigDecimal) map.get("QTY_PER_MOULD"));
            obj.setMoId((String) map.get("MO_ID"));
            obj.setInitialStock((BigDecimal) map.get("INITIAL_STOCK"));
            obj.setCapacity((BigDecimal) map.get("CAPACITY"));
            obj.setProductType((String) map.get("PRODUCT_TYPE"));
            obj.setPartNumber((BigDecimal) map.get("PART_NUMBER"));
            obj.setUpperConstant((BigDecimal) map.get("UPPER_CONSTANT"));
            obj.setMaxCapMonth0((BigDecimal) map.get("MAX_CAP_MONTH_0"));
            detailMarketingOrderListDualFrontRear.add(obj);
        }
	}
	    
	public void clearShift(BigDecimal partNum) {
	////System.out.println("masuk clear shift");
		List<String> listWct = new ArrayList<>(); // Gunakan ArrayList untuk penambahan elemen dinamis
		Iterator<ChangeMould> iterator = changeMouldList.iterator();
		while (iterator.hasNext()) {
		    ChangeMould cm = iterator.next();
		    if (cm.getPartNum().equals(partNum)) {
		        listWct.add(cm.getWct());
		        iterator.remove(); // Gunakan iterator untuk menghapus elemen dengan aman
		    }
		}
		
		for (int i = 0; i < changeMouldList.size(); i++) {
		    ChangeMould cm = changeMouldList.get(i);
		    if (cm.getPartNum().equals(partNum)) {
		        listWct.add(cm.getWct());
		        changeMouldList.remove(i);
		        i--; // Kurangi indeks karena elemen dihapus
		    }
		}
		
		for (int i = 0; i < endProductList.size(); i++) {
		    ChangeMould cm = endProductList.get(i);
		    if (cm.getPartNum().equals(partNum) && cm.getStatus() == 1) {
		        changeMouldList.remove(i);
		        i--; // Kurangi indeks karena elemen dihapus
		    }
		}
		
		machineCuringListTemp = machineCuringList;
		
		
		Iterator<ShiftMonthlyPlan> iteratorrr = newShiftMonthlyPlan.iterator();
		while (iterator.hasNext()) {
		    ShiftMonthlyPlan shf = iteratorrr.next();
		    if (shf.getPART_NUMBER().equals(partNum) && shf.getSTATUS().compareTo(BigDecimal.ONE) != 0) {
		    	order = order.add(shf.getTOTAL_KAPA());
		        iterator.remove(); // Gunakan iterator untuk menghapus elemen
		    }
		}
	}

    
    public void saveShift(List<ShiftMonthlyPlan> newShiftList) {
    	for(ShiftMonthlyPlan sf : newShiftList) {
    		shiftMonthlyRepo.save(sf);
    	}
    }
    
    public List<ShiftMonthlyPlan> flowTwentyFive(List<ShiftMonthlyPlan> sfList, List<MachineCuring> machineCuringList){
    	
    	
    	return sfList;
    }
    
    public List<CTCuring> getMachine(int status, String itemCuring) {
        List<CTCuring> ctCuringList = new ArrayList<>();

        if (status == 1) {
            ctCuringList = ctCuringRepo.findCBuildingMachineByItemCuring(itemCuring);
        } else if (status == 2) {
            ctCuringList = ctCuringRepo.findMachineByGCDBuilding(itemCuring);
        } else if (status == 3) {
            ctCuringList = ctCuringRepo.findMachineByABBuilding(itemCuring);
        } else if(status == 4) {
            ctCuringList = ctCuringRepo.findMachineByHBuilding(itemCuring);
        } else if(status == 12) { 
            ctCuringList = ctCuringRepo.findMachineABByGCDBuilding(itemCuring);
        } else if(status == 13) { 
            ctCuringList = ctCuringRepo.findMachineABByABBuilding(itemCuring);
        } else if(status == 14) { 
            ctCuringList = ctCuringRepo.findMachineABByHBuilding(itemCuring);
        } else if(status == 15) { // ALL A/B
        	ctCuringList = ctCuringRepo.findMachineABAllBuilding(itemCuring);
        } else if(status == 22) { 
            ctCuringList = ctCuringRepo.findMachineBOMByGCDBuilding(itemCuring);
        } else if(status == 23) { 
            ctCuringList = ctCuringRepo.findMachineBOMByABBuilding(itemCuring);
        } else if(status == 24) { 
            ctCuringList = ctCuringRepo.findMachineBOMByHBuilding(itemCuring);
        } else if(status == 25) { // ALL A/B
            ctCuringList = ctCuringRepo.findMachineBOMAllBuilding(itemCuring);
        } else if(status == 32) { 
            ctCuringList = ctCuringRepo.findMachine2CavByGCDBuilding(itemCuring);
        } else if(status == 33) { 
            ctCuringList = ctCuringRepo.findMachine2CavByABBuilding(itemCuring);
        } else if(status == 34) { 
            ctCuringList = ctCuringRepo.findMachine2CavByHBuilding(itemCuring);
        } else if(status == 35) { // ALL 2CAV
        	ctCuringList = ctCuringRepo.findMachine2CavAllBuilding(itemCuring);
        } else if(status == 99) {
        	ctCuringList = ctCuringRepo.findMachineByItemCuring(itemCuring);
        }

        return ctCuringList;
    }
    
    public int getStatusPrioritasMesin(DetailMo dtMo, BigDecimal smallOrderLimit) {
    	if ("TT".equals(dtMo.getProductType()) && dtMo.getRim().compareTo(BigDecimal.valueOf(14)) == 0) { // flowchart 11
            ////System.out.println("Done flow 11");
            return 1;
        } else if ("TT".equals(dtMo.getProductType())) { // Flowchart 14
            ////System.out.println("Done flow 14");
            return 12;
        } else if ("SINGLE COMPOUND".equals(dtMo.getExtDescription())) { // Flowchart 12
            ////System.out.println("Done flow 12");
            return 2;
        } else if ("TL".equals(dtMo.getProductType())) { // Flowchart 13
            ////System.out.println("Done flow 13");
            return 22;
        }
    	
        if (dtMo.getTotalAR().compareTo(smallOrderLimit) <= 0) { // flowchart 15
            ////System.out.println("Done flow 14");
            return 32;
        }
    	return 0;
    }
    
    public Date parseDate(String dateString) {
        Date date = null;
        try {
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
            date = dateFormat.parse(dateString); 
        } catch (Exception e) {
            e.printStackTrace();  
        }
        ////System.out.println("ini date " + date);
        return date;
    }
    
    public Date normalizeDate(Date date) {
        if (date == null) {
            return null; // Kembalikan null jika input null
        }
        
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.set(Calendar.HOUR_OF_DAY, 0);
        calendar.set(Calendar.MINUTE, 0);
        calendar.set(Calendar.SECOND, 0);
        calendar.set(Calendar.MILLISECOND, 0);
        
        return calendar.getTime(); // Kembalikan tanggal yang dinormalisasi
    }
    	
    public String getDayNameFromDate(Date date) {
        if (date == null) {
            return null; // Kembalikan null jika input null
        }
        SimpleDateFormat dayFormat = new SimpleDateFormat("EEEE", Locale.getDefault());
       ////System.out.println("hari " + dayFormat.format(date));
        return dayFormat.format(date); // Kembalikan nama hari dalam string
    }
    
    public Date getNextDay(Date date) {
        // Inisialisasi Calendar dengan tanggal yang diberikan
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);

        // Tambah satu hari
        calendar.add(Calendar.DAY_OF_YEAR, 1);

        // Mengembalikan tanggal berikutnya
        return calendar.getTime();
    }
    
    public Date getPreviousDay(Date date) {
        // Inisialisasi Calendar dengan tanggal yang diberikan
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);

        // Kurangi satu hari
        calendar.add(Calendar.DAY_OF_YEAR, -1);

        // Mengembalikan tanggal sebelumnya
        return calendar.getTime();
    }
    
    public String formatDateToString(Date date) {
        SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
        return formatter.format(date);
    }
    
    class PartWct{
    	private String wct;
    	private BigDecimal partNum;
		public String getWct() {
			return wct;
		}
		public void setWct(String wct) {
			this.wct = wct;
		}
		public BigDecimal getPartNum() {
			return partNum;
		}
		public void setPartNum(BigDecimal partNum) {
			this.partNum = partNum;
		}
    }
    
    public class ChangeMould{
    	private BigDecimal partNum;
    	private String wct;
    	private Date changeDate;
    	private int shift;
    	private int status;
    	
		public int getStatus() {
			return status;
		}
		public void setStatus(int status) {
			this.status = status;
		}
		public BigDecimal getPartNum() {
			return partNum;
		}
		public void setPartNum(BigDecimal partNum) {
			this.partNum = partNum;
		}
		public String getWct() {
			return wct;
		}
		public void setWct(String wct) {
			this.wct = wct;
		}
		public Date getChangeDate() {
			return changeDate;
		}
		public void setChangeDate(Date changeDate) {
			this.changeDate = changeDate;
		}
		public int getShift() {
			return shift;
		}
		public void setShift(int shift) {
			this.shift = shift;
		}
    	
    }
    
    class TempChangeMould{
    	private Date changeDate;
    	private int shift;
    	private int total;
		public Date getChangeDate() {
			return changeDate;
		}
		public void setChangeDate(Date changeDate) {
			this.changeDate = changeDate;
		}
		public int getShift() {
			return shift;
		}
		public void setShift(int shift) {
			this.shift = shift;
		}
		public int getTotal() {
			return total;
		}
		public void setTotal(int total) {
			this.total = total;
		}
    }
    
    class TempOrder{
    	private BigDecimal pratNum;
    	private BigDecimal marketingOrder;
    	private String itemCur;
		public TempOrder() {
		}
		
		public TempOrder(BigDecimal pratNum, BigDecimal marketingOrder, String itemCur) {
			this.pratNum = pratNum;
			this.marketingOrder = marketingOrder;
			this.itemCur = itemCur;
		}

		public BigDecimal getPratNum() {
			return pratNum;
		}
		public void setPratNum(BigDecimal pratNum) {
			this.pratNum = pratNum;
		}
		public BigDecimal getMarketingOrder() {
			return marketingOrder;
		}
		public void setMarketingOrder(BigDecimal marketingOrder) {
			this.marketingOrder = marketingOrder;
		}
		public String getItemCur() {
			return itemCur;
		}
		public void setItemCur(String itemCur) {
			this.itemCur = itemCur;
		}
    	
    }
    
    public class DetailMo {
        private BigDecimal ppd;
        private BigDecimal lowerConstant;
        private BigDecimal moMonth0;
        private BigDecimal totalAR;
        private String productCategory;
        private String machineType;
        private String itemExt;
        private String extDescription;
        private BigDecimal minOrder;
        private String sizeId;
        private BigDecimal detailId;
        private BigDecimal rim;
        private String itemCuring;
        private String wibTube;
        private String patternId;
        private String itemAssy;
        private String category;
        private BigDecimal cav;
        private BigDecimal qtyPerRak;
        private String description;
        private BigDecimal qtyPerMould;
        private String moId;
        private BigDecimal initialStock;
        private BigDecimal capacity;
        private String productType;
        private BigDecimal partNumber;
        private BigDecimal upperConstant;
        private BigDecimal maxCapMonth0;
        private BigDecimal productionLimit;

        // Getter dan Setter untuk semua properti
        public BigDecimal getPpd() { return ppd; }
        public void setPpd(BigDecimal ppd) { this.ppd = ppd; }
        public BigDecimal getLowerConstant() { return lowerConstant; }
        public void setLowerConstant(BigDecimal lowerConstant) { this.lowerConstant = lowerConstant; }
        public BigDecimal getMoMonth0() { return moMonth0; }
        public void setMoMonth0(BigDecimal moMonth0) { this.moMonth0 = moMonth0; }
        public String getProductCategory() { return productCategory; }
        public void setProductCategory(String productCategory) { this.productCategory = productCategory; }
        public String getMachineType() { return machineType; }
        public void setMachineType(String machineType) { this.machineType = machineType; }
        public String getItemExt() { return itemExt; }
        public void setItemExt(String itemExt) { this.itemExt = itemExt; }
        public String getExtDescription() { return extDescription; }
        public void setExtDescription(String extDescription) { this.extDescription = extDescription; }
        public BigDecimal getMinOrder() { return minOrder; }
        public void setMinOrder(BigDecimal minOrder) { this.minOrder = minOrder; }
        public String getSizeId() { return sizeId; }
        public void setSizeId(String sizeId) { this.sizeId = sizeId; }
        public BigDecimal getDetailId() { return detailId; }
        public void setDetailId(BigDecimal detailId) { this.detailId = detailId; }
        public BigDecimal getRim() { return rim; }
        public void setRim(BigDecimal rim) { this.rim = rim; }
        public String getItemCuring() { return itemCuring; }
        public void setItemCuring(String itemCuring) { this.itemCuring = itemCuring; }
        public String getWibTube() { return wibTube; }
        public void setWibTube(String wibTube) { this.wibTube = wibTube; }
        public String getPatternId() { return patternId; }
        public void setPatternId(String patternId) { this.patternId = patternId; }
        public String getItemAssy() { return itemAssy; }
        public void setItemAssy(String itemAssy) { this.itemAssy = itemAssy; }
        public String getCategory() { return category; }
        public void setCategory(String category) { this.category = category; }
        public BigDecimal getCav() { return cav; }
        public void setCav(BigDecimal cav) { this.cav = cav; }
        public BigDecimal getQtyPerRak() { return qtyPerRak; }
        public void setQtyPerRak(BigDecimal qtyPerRak) { this.qtyPerRak = qtyPerRak; }
        public String getDescription() { return description; }
        public void setDescription(String description) { this.description = description; }
        public BigDecimal getQtyPerMould() { return qtyPerMould; }
        public void setQtyPerMould(BigDecimal qtyPerMould) { this.qtyPerMould = qtyPerMould; }
        public String getMoId() { return moId; }
        public void setMoId(String moId) { this.moId = moId; }
        public BigDecimal getInitialStock() { return initialStock; }
        public void setInitialStock(BigDecimal initialStock) { this.initialStock = initialStock; }
        public BigDecimal getCapacity() { return capacity; }
        public void setCapacity(BigDecimal capacity) { this.capacity = capacity; }
        public String getProductType() { return productType; }
        public void setProductType(String productType) { this.productType = productType; }
        public BigDecimal getPartNumber() { return partNumber; }
        public void setPartNumber(BigDecimal partNumber) { this.partNumber = partNumber; }
        public BigDecimal getUpperConstant() { return upperConstant; }
        public void setUpperConstant(BigDecimal upperConstant) { this.upperConstant = upperConstant; }
        public BigDecimal getMaxCapMonth0() { return maxCapMonth0; }
        public void setMaxCapMonth0(BigDecimal maxCapMonth0) { this.maxCapMonth0 = maxCapMonth0; }
        
        public BigDecimal getTotalAR() {
			return totalAR;
		}
		public void setTotalAR(BigDecimal totalAR) {
			this.totalAR = totalAR;
		}
		public BigDecimal getProductionLimit() {
			return productionLimit;
		}
		public void setProductionLimit(BigDecimal productionLimit) {
			this.productionLimit = productionLimit;
		}
		@Override
        public String toString() {
            return "CustomObject{" +
                    "ppd=" + ppd +
                    ", lowerConstant='" + lowerConstant + '\'' +
                    ", moMonth0=" + moMonth0 +
                    ", productCategory='" + productCategory + '\'' +
                    ", machineType='" + machineType + '\'' +
                    ", itemExt='" + itemExt + '\'' +
                    ", extDescription='" + extDescription + '\'' +
                    ", minOrder=" + minOrder +
                    ", sizeId='" + sizeId + '\'' +
                    ", detailId=" + detailId +
                    ", rim=" + rim +
                    ", itemCuring='" + itemCuring + '\'' +
                    ", wibTube='" + wibTube + '\'' +
                    ", patternId='" + patternId + '\'' +
                    ", itemAssy='" + itemAssy + '\'' +
                    ", category='" + category + '\'' +
                    ", cav=" + cav +
                    ", qtyPerRak=" + qtyPerRak +
                    ", description='" + description + '\'' +
                    ", qtyPerMould=" + qtyPerMould +
                    ", moId='" + moId + '\'' +
                    ", initialStock=" + initialStock +
                    ", capacity=" + capacity +
                    ", productType='" + productType + '\'' +
                    ", partNumber='" + partNumber + '\'' +
                    ", upperConstant='" + upperConstant + '\'' +
                    ", totalAR='" + totalAR + '\'' +
                    ", maxCapMonth0=" + maxCapMonth0 +
                    '}';
        }
    }
    
    public ByteArrayInputStream exportExcel(int month, int year, int limitChange, BigDecimal minA, BigDecimal maxA, BigDecimal minB, BigDecimal maxB, BigDecimal minC, BigDecimal maxC, BigDecimal minD, BigDecimal maxD) throws IOException {
    	List<ShiftMonthlyPlan> shiftMonthlyPlan = MonthlyPlan(month, year, limitChange, minA, maxA, minB, maxB, minC, maxC, minD, maxD);
    	System.out.println(shiftMonthlyPlan.size());
    	List<String> productDescription = new ArrayList<>();
    	
    		for (int i = 0; i < shiftMonthlyPlan.size(); i++) {
    		    BigDecimal partNumber = shiftMonthlyPlan.get(i).getPART_NUMBER();
    		    String description = shiftMonthlyRepo.findDescriptionByPartNum(partNumber);
    		    if (description == null) {
    		        description = "N/A";
    		    }
    		    productDescription.add(description);
    		}


        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
        	//style
        	// Font
            Font calibri11 = workbook.createFont();
            calibri11.setFontName("Calibri");
            calibri11.setFontHeightInPoints((short) 11);
            
            Font calibri12 = workbook.createFont();
            calibri12.setFontName("Calibri");
            calibri12.setFontHeightInPoints((short) 12);
            
            Font calibriBold11 = workbook.createFont();
            calibriBold11.setFontName("Calibri");
            calibriBold11.setFontHeightInPoints((short) 11);
            calibriBold11.setBold(true);
            
            Font calibriBold12 = workbook.createFont();
            calibriBold12.setFontName("Calibri");
            calibriBold12.setFontHeightInPoints((short) 12);
            calibriBold12.setBold(true);
            
            Font calibriBold14 = workbook.createFont();
            calibriBold14.setFontName("Calibri");
            calibriBold14.setFontHeightInPoints((short) 14);
            calibriBold14.setBold(true);
            // End Font
            
            // Border cell style
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            // End border cell style

            // Style untuk cell
            CellStyle title = workbook.createCellStyle();
            title.setFont(calibriBold14);
            title.setAlignment(HorizontalAlignment.CENTER);
            title.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri11Left = workbook.createCellStyle();
            calibri11Left.setFont(calibri11);
            calibri11Left.setAlignment(HorizontalAlignment.LEFT);
            calibri11Left.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri11Right = workbook.createCellStyle();
            calibri11Right.setFont(calibri11);
            calibri11Right.setAlignment(HorizontalAlignment.RIGHT);
            calibri11Right.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri11Center = workbook.createCellStyle();
            calibri11Center.setFont(calibri11);
            calibri11Center.setAlignment(HorizontalAlignment.CENTER);
            calibri11Center.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri11LeftBorder = workbook.createCellStyle();
            calibri11LeftBorder.cloneStyleFrom(borderStyle);
            calibri11LeftBorder.setFont(calibri11);
            calibri11LeftBorder.setAlignment(HorizontalAlignment.LEFT);
            calibri11LeftBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibriBold11LeftBorder = workbook.createCellStyle();
            calibriBold11LeftBorder.cloneStyleFrom(borderStyle);
            calibriBold11LeftBorder.setFont(calibriBold11);
            calibriBold11LeftBorder.setAlignment(HorizontalAlignment.LEFT);
            calibriBold11LeftBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri11RightBorder = workbook.createCellStyle();
            calibri11RightBorder.cloneStyleFrom(borderStyle);
            calibri11RightBorder.setFont(calibri11);
            calibri11RightBorder.setAlignment(HorizontalAlignment.RIGHT);
            calibri11RightBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri11CenterBorder = workbook.createCellStyle();
            calibri11CenterBorder.cloneStyleFrom(borderStyle);
            calibri11CenterBorder.setFont(calibri11);
            calibri11CenterBorder.setAlignment(HorizontalAlignment.CENTER);
            calibri11CenterBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibriBold11CenterBorder = workbook.createCellStyle();
            calibriBold11CenterBorder.cloneStyleFrom(borderStyle);
            calibriBold11CenterBorder.setFont(calibriBold11);
            calibriBold11CenterBorder.setAlignment(HorizontalAlignment.CENTER);
            calibriBold11CenterBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CreationHelper createHelper = workbook.getCreationHelper();
            CellStyle calibri11Date = workbook.createCellStyle();
            calibri11Date.cloneStyleFrom(borderStyle);
            calibri11Date.setFont(calibri11);
            calibri11Date.setAlignment(HorizontalAlignment.CENTER);
            calibri11Date.setVerticalAlignment(VerticalAlignment.CENTER);
            calibri11Date.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
            //end style
            
            //mesin curing sheet
            Sheet curingSheet = workbook.createSheet("mesin curing");
            curingSheet.setColumnWidth(1, 3000);
            curingSheet.setColumnWidth(4, 10000);
            curingSheet.setColumnWidth(5, 3000);
            curingSheet.setColumnWidth(6, 10000);
            
            //header table
            Row tableHeadCuringRow = curingSheet.createRow(1);
            Cell tableHeadCuringCell;
            
            String[] curingHeaderLabels = {"Tanggal", "Nomor", "Cavity", "Work Center Text", "Item Curing", "Deskripsi", "Shift 1", "Shift 2", "Shift 3", "Total"};
            CellStyle[] curingHeaderStyles = {calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, 
            		calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder};

            // Loop untuk kolom 0-5
            for (int col = 0; col < curingHeaderLabels.length; col++) {
            	tableHeadCuringCell = tableHeadCuringRow.createCell(col+1);
            	tableHeadCuringCell.setCellStyle(curingHeaderStyles[col]);
                if (!curingHeaderLabels[col].isEmpty()) {
                	tableHeadCuringCell.setCellValue(curingHeaderLabels[col]);
                }
            }
            
            //data tables
            
            int curingDatarow = 2;
            Row curingDataRow;
            Cell curingDataCell;
            for (int j = 0; j < shiftMonthlyPlan.size(); j++) {
            	
            	curingDataRow = curingSheet.createRow(curingDatarow);
            	curingDataCell = curingDataRow.createCell(1);
                curingDataCell.setCellStyle(calibri11Date);
                curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getDATE());
                
                
                curingDataCell = curingDataRow.createCell(2);
                curingDataCell.setCellStyle(calibriBold11CenterBorder);
                curingDataCell.setCellValue(j+1);
                
                curingDataCell = curingDataRow.createCell(3);
        		curingDataCell.setCellStyle(calibri11CenterBorder);
        		curingDataCell.setCellValue("C");
        		
        		curingDataCell = curingDataRow.createCell(4);
        		curingDataCell.setCellStyle(calibriBold11LeftBorder);
        		curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getWORK_CENTER_TEXT());
        		
        		curingDataCell = curingDataRow.createCell(5);
        		curingDataCell.setCellStyle(calibriBold11LeftBorder);
        		curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getITEM_CURING());
        		
        		curingDataCell = curingDataRow.createCell(6);
        		curingDataCell.setCellStyle(calibri11LeftBorder);
        		curingDataCell.setCellValue(productDescription.get(j));
        		
        		curingDataCell = curingDataRow.createCell(7);
        		curingDataCell.setCellStyle(calibri11RightBorder);
        		curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getKAPA_SHIFT_1().doubleValue());
        		
        		curingDataCell = curingDataRow.createCell(8);
        		curingDataCell.setCellStyle(calibri11RightBorder);
        		curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getKAPA_SHIFT_2().doubleValue());
        		
        		curingDataCell = curingDataRow.createCell(9);
        		curingDataCell.setCellStyle(calibri11RightBorder);
        		curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getKAPA_SHIFT_3().doubleValue());
        		
        		curingDataCell = curingDataRow.createCell(10);
        		curingDataCell.setCellStyle(calibri11RightBorder);
        		curingDataCell.setCellValue(shiftMonthlyPlan.get(j).getTOTAL_KAPA().doubleValue());
                
                curingDatarow++;
            }
            //end curing sheet
        	
            //prepare prod sheet
            Sheet prepareProdSheet = workbook.createSheet("PREPARE PRODE TIRE");
            
            // Set column width
            prepareProdSheet.setColumnWidth(1, 5000);
            prepareProdSheet.setColumnWidth(4, 8000);
            prepareProdSheet.setColumnWidth(5, 8000);

            //header
            Row monthRow = prepareProdSheet.createRow(8);
            Cell monthCell = monthRow.createCell(4);
            monthCell.setCellStyle(calibri11Right);
            monthCell.setCellValue("MONTH OF :");
            
            Row sectionRow = prepareProdSheet.createRow(9);
            Cell sectionCell = sectionRow.createCell(4);
            sectionCell.setCellStyle(calibri11Right);
            sectionCell.setCellValue("SECTION :");
            
            Row issueDateRow = prepareProdSheet.createRow(11);
            Cell issueDateCell = issueDateRow.createCell(4);
            issueDateCell.setCellStyle(calibri11Right);
            issueDateCell.setCellValue("ISSUE DATE :");
            
            prepareProdSheet.addMergedRegion(new CellRangeAddress(8, 9, 18, 23));
            Cell titleCell = monthRow.createCell(18);
            titleCell.setCellStyle(title);
            titleCell.setCellValue("M O N T H L Y      P L A N N I N G");
            
            Cell kadeptCell = monthRow.createCell(41);
            kadeptCell.setCellStyle(calibri11CenterBorder);
            kadeptCell.setCellValue("KADEPT");
            
            Cell kassieCell = monthRow.createCell(42);
            kassieCell.setCellStyle(calibri11CenterBorder);
            kassieCell.setCellValue("KASSIE PP");
            
            Cell docNumCell = monthRow.createCell(43);
            docNumCell.setCellStyle(calibri11LeftBorder);
            docNumCell.setCellValue("NO. DOK");
            
            //table head
            int i;
            Row tableHeadMpRow1 = prepareProdSheet.createRow(16);
            Row tableHeadMpRow2 = prepareProdSheet.createRow(17);
            Cell tableHeadMpCell;

            // Array untuk header teks dan gaya untuk kolom 0-5
            String[] headerMpLabels = {"", "PART NUMBER", "", "NO.", "SIZE", "PATTERN"};
            CellStyle[] headerMpStyles = {calibri11CenterBorder, calibriBold11CenterBorder, calibri11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder};

            // Loop untuk kolom 0-5
            for (int col = 0; col < headerMpLabels.length; col++) {
            	prepareProdSheet.addMergedRegion(new CellRangeAddress(16, 17, col, col));
                tableHeadMpCell = tableHeadMpRow1.createCell(col);
                tableHeadMpCell.setCellStyle(headerMpStyles[col]);
                if (!headerMpLabels[col].isEmpty()) {
                    tableHeadMpCell.setCellValue(headerMpLabels[col]);
                }
                tableHeadMpCell = tableHeadMpRow2.createCell(col);
                tableHeadMpCell.setCellStyle(calibri11CenterBorder);
            }

            // Mendapatkan jumlah hari di bulan ini
            YearMonth yearMonth = YearMonth.of(2024, 11);
            int jumlahHariBulanIni = yearMonth.lengthOfMonth();
            LocalDate localDate;
            Date date = new Date();

            // Loop untuk menambah tanggal (kolom 6 ke atas)
            for (i = 0; i < jumlahHariBulanIni; i++) {
                int col = i + 6;
                prepareProdSheet.addMergedRegion(new CellRangeAddress(16, 17, col, col));
                tableHeadMpCell = tableHeadMpRow1.createCell(col);
                tableHeadMpCell.setCellStyle(calibriBold11CenterBorder);
                tableHeadMpCell.setCellValue(i + 1);
                tableHeadMpCell = tableHeadMpRow2.createCell(col);
                tableHeadMpCell.setCellStyle(calibri11CenterBorder);
            }

            // Menambahkan kolom "TOTAL" setelah jumlah hari dalam bulan
            int totalCol = i + 6;
            tableHeadMpCell = tableHeadMpRow1.createCell(totalCol);
            tableHeadMpCell.setCellStyle(calibriBold11CenterBorder);
            tableHeadMpCell.setCellValue("TOTAL");
            tableHeadMpCell = tableHeadMpRow2.createCell(totalCol);
            tableHeadMpCell.setCellStyle(calibri11CenterBorder);
            tableHeadMpCell.setCellValue(i);
            
            int mpDatarow = 18;
            Row mpDataRow;
            Cell mpDataCell;
            Set<String> addedPartNumbers = new HashSet<>();
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");

            for (int j = 0; j < shiftMonthlyPlan.size(); j++) {
                String partNumber = String.valueOf(shiftMonthlyPlan.get(j).getPART_NUMBER());
                String description = productDescription.get(j);
                int intCapacity = 0;
                int totalCapacity = 0;
                String planDate = "";
                String actualDate = "";

                
                // Jika partNumber belum pernah ditambahkan
                if (partNumber != null && !addedPartNumbers.contains(partNumber)) {
                    // Tambahkan ke HashSet untuk melacak
                    addedPartNumbers.add(partNumber);

                    // Buat baris dan tambahkan data
                    mpDataRow = prepareProdSheet.createRow(mpDatarow);
                    mpDataCell = mpDataRow.createCell(0);
                    //mpDataCell.setCellValue("");
                    mpDataCell.setCellStyle(calibri11LeftBorder);
                    
                    mpDataCell = mpDataRow.createCell(1);
                    mpDataCell.setCellValue(partNumber);
                    mpDataCell.setCellStyle(calibri11CenterBorder);
                    
                    mpDataCell = mpDataRow.createCell(2);
                    //mpDataCell.setCellValue("");
                    mpDataCell.setCellStyle(calibri11LeftBorder);
                    
                    mpDataCell = mpDataRow.createCell(3);
                    mpDataCell.setCellValue(mpDatarow-17);
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    mpDataCell = mpDataRow.createCell(4);
                    mpDataCell.setCellValue(description);
                    mpDataCell.setCellStyle(calibri11LeftBorder);
                    
                    mpDataCell = mpDataRow.createCell(5);
                    mpDataCell.setCellValue(description);
                    mpDataCell.setCellStyle(calibri11LeftBorder);
                    
                    //cap day 1
                    intCapacity = 0;
                    localDate = yearMonth.atDay(1);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
                    	actualDate = dateFormat.format(date);
                        if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
                           ////System.out.println(planDate);
                           ////System.out.println(actualDate);
                        	intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 1 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(6);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 2
                    intCapacity = 0;
                    localDate = yearMonth.atDay(2);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 2 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(7);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 3
                    intCapacity = 0;
                    localDate = yearMonth.atDay(3);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 3 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(8);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                  //cap day 4
                    intCapacity = 0;
                    localDate = yearMonth.atDay(4);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 4 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(9);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 5
                    intCapacity = 0;
                    localDate = yearMonth.atDay(5);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 5 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(10);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 6
                    intCapacity = 0;
                    localDate = yearMonth.atDay(6);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 6 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(11);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 7
                    intCapacity = 0;
                    localDate = yearMonth.atDay(7);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 7 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(12);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 8
                    intCapacity = 0;
                    localDate = yearMonth.atDay(8);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 8 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(13);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 9
                    intCapacity = 0;
                    localDate = yearMonth.atDay(9);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 9 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(14);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 10
                    intCapacity = 0;
                    localDate = yearMonth.atDay(10);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 10 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(15);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 11
                    intCapacity = 0;
                    localDate = yearMonth.atDay(11);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 11 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(16);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 12
                    intCapacity = 0;
                    localDate = yearMonth.atDay(12);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 12 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(17);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 13
                    intCapacity = 0;
                    localDate = yearMonth.atDay(13);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 13 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(18);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 14
                    intCapacity = 0;
                    localDate = yearMonth.atDay(14);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 14 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(19);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 15
                    intCapacity = 0;
                    localDate = yearMonth.atDay(15);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 15 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(20);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 16
                    intCapacity = 0;
                    localDate = yearMonth.atDay(16);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 16 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(21);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 17
                    intCapacity = 0;
                    localDate = yearMonth.atDay(17);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 17 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(22);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 18
                    intCapacity = 0;
                    localDate = yearMonth.atDay(18);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 18 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(23);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 19
                    intCapacity = 0;
                    localDate = yearMonth.atDay(19);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 19 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(24);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 20
                    intCapacity = 0;
                    localDate = yearMonth.atDay(20);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 20 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(25);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 21
                    intCapacity = 0;
                    localDate = yearMonth.atDay(21);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 21 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(26);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 22
                    intCapacity = 0;
                    localDate = yearMonth.atDay(22);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 22 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(27);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 23
                    intCapacity = 0;
                    localDate = yearMonth.atDay(23);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 23 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(28);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 24
                    intCapacity = 0;
                    localDate = yearMonth.atDay(24);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 24 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(29);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 25
                    intCapacity = 0;
                    localDate = yearMonth.atDay(25);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 25 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(30);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 26
                    intCapacity = 0;
                    localDate = yearMonth.atDay(26);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 26 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(31);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 27
                    intCapacity = 0;
                    localDate = yearMonth.atDay(27);
                    planDate = "";
                	actualDate = "";
                    date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                    for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
	                    	actualDate = dateFormat.format(date);
	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
	                           ////System.out.println(planDate);
	                           ////System.out.println(actualDate);
                            intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                        }
                    }
                   ////System.out.println(partNumber + " at day 27 : " + intCapacity);
                    totalCapacity += intCapacity;
                    
                    mpDataCell = mpDataRow.createCell(32);
                    if(intCapacity > 0) {
                    	mpDataCell.setCellValue((double) intCapacity);
                    } else {
                    	mpDataCell.setCellValue("");
                    }
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    //cap day 28
                    if (yearMonth.lengthOfMonth() >= 28) {
                    	intCapacity = 0;
                        localDate = yearMonth.atDay(28);
                        planDate = "";
                    	actualDate = "";
                        date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                        for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
    	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
    	                    	actualDate = dateFormat.format(date);
    	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
    	                           ////System.out.println(planDate);
    	                           ////System.out.println(actualDate);
                                intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                            }
                        }
                       ////System.out.println(partNumber + " at day 28 : " + intCapacity);
                        totalCapacity += intCapacity;
                        
                        mpDataCell = mpDataRow.createCell(33);
                        if(intCapacity > 0) {
                        	mpDataCell.setCellValue((double) intCapacity);
                        } else {
                        	mpDataCell.setCellValue("");
                        }
                        mpDataCell.setCellStyle(calibri11RightBorder);
                    }
                    
                    //cap day 29
                    if (yearMonth.lengthOfMonth() >= 29) {
                    	intCapacity = 0;
                        localDate = yearMonth.atDay(29);
                        planDate = "";
                    	actualDate = "";
                        date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                        for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
    	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
    	                    	actualDate = dateFormat.format(date);
    	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
    	                           ////System.out.println(planDate);
    	                           ////System.out.println(actualDate);
                                intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                            }
                        }
                       ////System.out.println(partNumber + " at day 29 : " + intCapacity);
                        totalCapacity += intCapacity;
                        
                        mpDataCell = mpDataRow.createCell(34);
                        if(intCapacity > 0) {
                        	mpDataCell.setCellValue((double) intCapacity);
                        } else {
                        	mpDataCell.setCellValue("");
                        }
                        mpDataCell.setCellStyle(calibri11RightBorder);
                    }
                    
                    //cap day 30
                    if (yearMonth.lengthOfMonth() >= 30) {
                    	intCapacity = 0;
                        localDate = yearMonth.atDay(30);
                        planDate = "";
                    	actualDate = "";
                        date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                        for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
    	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
    	                    	actualDate = dateFormat.format(date);
    	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
    	                           ////System.out.println(planDate);
    	                           ////System.out.println(actualDate);
                                intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                            }
                        }
                       ////System.out.println(partNumber + " at day 30 : " + intCapacity);
                        totalCapacity += intCapacity;
                        
                        mpDataCell = mpDataRow.createCell(35);
                        if(intCapacity > 0) {
                        	mpDataCell.setCellValue((double) intCapacity);
                        } else {
                        	mpDataCell.setCellValue("");
                        }
                        mpDataCell.setCellStyle(calibri11RightBorder);
                    }
                    
                    //cap day 31
                    if (yearMonth.lengthOfMonth() >= 31) {
                    	intCapacity = 0;
                        localDate = yearMonth.atDay(31);
                        planDate = "";
                    	actualDate = "";
                        date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                        for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
    	                    	planDate = dateFormat.format(shiftMonthlyPlan.get(k).getDATE());
    	                    	actualDate = dateFormat.format(date);
    	                    	if (partNumber.equals(String.valueOf(shiftMonthlyPlan.get(k).getPART_NUMBER())) && planDate.equals(actualDate)) {
    	                           ////System.out.println(planDate);
    	                           ////System.out.println(actualDate);
                                intCapacity = intCapacity + shiftMonthlyPlan.get(k).getKAPA_SHIFT_1().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_2().intValue() + shiftMonthlyPlan.get(k).getKAPA_SHIFT_3().intValue(); 
                            }
                        }
                       ////System.out.println(partNumber + " at day 31 : " + intCapacity);
                        totalCapacity += intCapacity;
                        
                        mpDataCell = mpDataRow.createCell(36);
                        if(intCapacity > 0) {
                        	mpDataCell.setCellValue((double) intCapacity);
                        } else {
                        	mpDataCell.setCellValue("");
                        }
                        mpDataCell.setCellStyle(calibri11RightBorder);
                    }
                    
                    mpDataCell = mpDataRow.createCell(jumlahHariBulanIni + 6);
                    mpDataCell.setCellValue((double) totalCapacity);
                    mpDataCell.setCellStyle(calibri11RightBorder);
                    
                    totalCapacity = 0;
                    mpDatarow++; // Naikkan index baris
                }
            }
            //end prepare prod sheet
            
            //change mould sheet
            Sheet changeMouldSheet = workbook.createSheet("CHANGE MOULD");
            
            changeMouldSheet.setColumnWidth(3, 8000);
            changeMouldSheet.setColumnWidth(1, 5000);
            
            Row tableHeadChangeMouldRow = changeMouldSheet.createRow(1);
            Cell tableHeadChangeMouldCell;
            
            String[] changeMouldHeaderLabels = {"Tanggal", "Shift", "Work Center Text", "Part Number"};
            CellStyle[] changeMouldHeaderStyles = {calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder, calibriBold11CenterBorder};

            // Loop untuk kolom 0-5
            for (int col = 0; col < changeMouldHeaderLabels.length; col++) {
            	tableHeadChangeMouldCell = tableHeadChangeMouldRow.createCell(col+1);
            	tableHeadChangeMouldCell.setCellStyle(changeMouldHeaderStyles[col]);
                if (!changeMouldHeaderLabels[col].isEmpty()) {
                	tableHeadChangeMouldCell.setCellValue(changeMouldHeaderLabels[col]);
                }
            }
            
            int changeMouldRow = 2;
            Row changeMouldDataRow;
            Cell changeMouldDataCell;
            
            for (int col = 0; col < changeMouldList.size(); col++) {
            	changeMouldDataRow = changeMouldSheet.createRow(changeMouldRow);
            	
            	changeMouldDataCell = changeMouldDataRow.createCell(1);
            	changeMouldDataCell.setCellStyle(calibri11Date);
            	changeMouldDataCell.setCellValue(changeMouldList.get(col).getChangeDate());
            	
            	changeMouldDataCell = changeMouldDataRow.createCell(2);
            	changeMouldDataCell.setCellStyle(calibriBold11CenterBorder);
            	changeMouldDataCell.setCellValue(changeMouldList.get(col).getShift());
            	
            	changeMouldDataCell = changeMouldDataRow.createCell(3);
            	changeMouldDataCell.setCellStyle(calibriBold11CenterBorder);
            	changeMouldDataCell.setCellValue(changeMouldList.get(col).getWct());
            	
            	changeMouldDataCell = changeMouldDataRow.createCell(4);
            	changeMouldDataCell.setCellStyle(calibriBold11CenterBorder);
            	changeMouldDataCell.setCellValue(changeMouldList.get(col).getPartNum().toString());
            	
            	changeMouldRow++;
            }
            //end change mould sheet
            
            workbook.write(out); // Menulis data ke output stream
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
           ////System.out.println("Fail to export data");
            return null;
        } finally {
            out.close(); // Tutup output stream setelah selesai
        }
    }
    
    
    
    public ViewMonthlyPlanning getDetailMonthlyPlan(int month, int year, int limitChange, BigDecimal minA, BigDecimal maxA, BigDecimal minB, BigDecimal maxB, BigDecimal minC, BigDecimal maxC, BigDecimal minD, BigDecimal maxD) {
        List<ShiftMonthlyPlan> shiftMonthlyPlan = MonthlyPlan(month, year, limitChange, minA, maxA, minB, maxB, minC, maxC, minD, maxD);

        ViewMonthlyPlanning viewMonthlyPlanning = new ViewMonthlyPlanning();
        List<DetailMonthlyPlanCuring> detailList1 = new ArrayList<>();
        List<DetailDailyMonthlyPlanCuring> detailList2 = new ArrayList<>();
        List<Map<String, Object>> description = new ArrayList<>();
        List<Map<String, Object>> productDetails = new ArrayList<>();

        BigDecimal detailId = BigDecimal.ONE;
        BigDecimal detailDailyId = BigDecimal.ONE;
        
        Map<BigDecimal, BigDecimal> partNumberToDetailIdCuringMap = new HashMap<>();
        
        int totalKapa = 0;
        int totalPlanDaily = 0;

        // totalKapa per partNumber 
        for (int j = 0; j < shiftMonthlyPlan.size(); j++) {
            final BigDecimal partNumber = shiftMonthlyPlan.get(j).getPART_NUMBER(); 
            
            // Check if partNumber already exists in detailList1
            boolean isPartNumberExists = detailList1.stream()
                .anyMatch(detail -> detail.getPartNumber().equals(partNumber));
            
            if (isPartNumberExists) {
                continue; 
            }

            for (int k = 0; k < shiftMonthlyPlan.size(); k++) {
                if (partNumber.equals(shiftMonthlyPlan.get(k).getPART_NUMBER())) {
                    totalKapa += shiftMonthlyPlan.get(k).getTOTAL_KAPA().intValue(); 
                }
            }

            DetailMonthlyPlanCuring detail = new DetailMonthlyPlanCuring();
            detail.setDetailIdCuring(detailId);
            detail.setPartNumber(partNumber);
            detail.setTotal(BigDecimal.valueOf(totalKapa));
            
            detailId = detailId.add(BigDecimal.ONE);

            detailList1.add(detail);
            partNumberToDetailIdCuringMap.put(partNumber, detail.getDetailIdCuring());

            // Reset totalKapa 
            totalKapa = 0;
        }

        // totalPlanDaily for each partNumber per day 
        for (ShiftMonthlyPlan shift : shiftMonthlyPlan) {
            BigDecimal partNumber = shift.getPART_NUMBER();
            final Date planDate = shift.getDATE();
            
            BigDecimal detailIdCuring = partNumberToDetailIdCuringMap.get(partNumber);

            // Skip if the combination of detailIdCuring and planDate already exists in detailList2
            boolean isAlreadyProcessed = detailList2.stream()
                .anyMatch(detail -> detail.getDetailIdCuring().equals(detailIdCuring) && detail.getDateDailyMp().equals(planDate));
            
            if (isAlreadyProcessed) {
                continue; 
            }

            for (ShiftMonthlyPlan shiftInner : shiftMonthlyPlan) {
                if (partNumber.equals(shiftInner.getPART_NUMBER()) && planDate.equals(shiftInner.getDATE())) {
                    totalPlanDaily += shiftInner.getTOTAL_KAPA().intValue();
                }
            }

            DetailDailyMonthlyPlanCuring detail = new DetailDailyMonthlyPlanCuring();
            detail.setDetailDailyIdCuring(detailDailyId);
            detail.setDetailIdCuring(detailIdCuring);
            detail.setDateDailyMp(planDate);
            detail.setTotalPlan(BigDecimal.valueOf(totalPlanDaily));
            
            detailDailyId = detailDailyId.add(BigDecimal.ONE); 
            detailList2.add(detail);

            // Reset totalPlanDaily 
            totalPlanDaily = 0;
            BigDecimal partNum = shift.getPART_NUMBER();
            if (productDetails.stream().noneMatch(detail1 -> detail1.get("partNumber").equals(partNum))) {
                String description2 = shiftMonthlyRepo.findDescriptionByPartNum(partNum);
                Map<String, Object> detail2 = new HashMap<>();
                detail2.put("partNumber", partNum);
                detail2.put("description", description2 != null ? description2 : "N/A");

                productDetails.add(detail2);
            }
        }
        
        List<Map<String, Object>> descriptionWD = monthlyPlanningRepo.getDescriptionWD(month, year);

        viewMonthlyPlanning.setDetailMonthlyPlanCuring(detailList1);
        viewMonthlyPlanning.setDetailDailyMonthlyPlanCuring(detailList2);
        viewMonthlyPlanning.setShiftMonthlyPlan(shiftMonthlyPlan);
        viewMonthlyPlanning.setChangeMould(changeMouldList);
        viewMonthlyPlanning.setDescription(descriptionWD);
        viewMonthlyPlanning.setProductDetails(productDetails);
        
        System.out.println("Data shift: " + shiftMonthlyPlan.size());
        System.out.println("Data list1 (detailMonthlyPlanCuring): " + detailList1.size());
        System.out.println("Data list2 (detailDailyMonthlyPlanCuring): " + detailList2.size());

        return viewMonthlyPlanning;
    }

}