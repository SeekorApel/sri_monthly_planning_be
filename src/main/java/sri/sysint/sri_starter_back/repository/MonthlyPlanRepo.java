package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.DetailDailyMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.DetailMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.MonthlyPlan;
import sri.sysint.sri_starter_back.model.MonthlyPlanningCuring;

public interface MonthlyPlanRepo extends JpaRepository<MonthlyPlan, String>{
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_MONTHLYPLAN_CURING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_DAILYMONTHLYPLAN_CURING", nativeQuery = true)
    BigDecimal getNewIdDaily();
	
	@Query(value = "SELECT \r\n"
			+ "    p.PART_NUMBER,\r\n"
			+ "    p.ITEM_CURING,\r\n"
			+ "    i.KAPA_PER_MOULD,\r\n"
			+ "    i.MOULD_MONTHLY_PLAN,\r\n"
			+ "    p.ITEM_ASSY,\r\n"
			+ "    t.PRODUCT_TYPE,\r\n"
			+ "    p.EXT_DESCRIPTION\r\n"
			+ "FROM \r\n"
			+ "    SRI_IMPP_M_PRODUCT p\r\n"
			+ "JOIN \r\n"
			+ "    SRI_IMPP_M_ITEMCURING i \r\n"
			+ "ON \r\n"
			+ "    p.ITEM_CURING = i.ITEM_CURING\r\n"
			+ "JOIN \r\n"
			+ "    SRI_IMPP_M_PRODUCTTYPE t\r\n"
			+ "ON \r\n"
			+ "    p.PRODUCT_TYPE_ID = t.PRODUCT_TYPE_ID\r\n"
			+ "WHERE\r\n"
			+ "    p.STATUS = 1", nativeQuery = true)
	List<Map<String, Object>> getData();
	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MONTHLYPLAN_CURING ORDER BY MP_CURING_ID ASC", nativeQuery = true)
	List<MonthlyPlanningCuring> findAllMP();
	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MONTHLYPLAN_CURING WHERE MP_CURING_ID = :docNum ORDER BY MP_CURING_ID ASC", nativeQuery = true)
	MonthlyPlanningCuring findMpById(@Param ("docNum") String mpId);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_MONTHLYPLAN_CURING WHERE DOC_NUM_CURING = :mpId ORDER BY DOC_NUM_CURING ASC", nativeQuery = true)
	List<DetailMonthlyPlanCuring> findDetailMpById(@Param ("mpId") String mpId);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_DAILYMONTHLYPLAN_CURING WHERE DETAIL_ID_CURING = :detailId ORDER BY DETAIL_ID_CURING ASC", nativeQuery = true)
	List<DetailDailyMonthlyPlanCuring> findDetailDailyMpById(@Param ("detailId") BigDecimal detailId);
	
	@Query(value = "SELECT \r\n"
			+ "    mp.DETAIL_ID_CURING,\r\n"
			+ "    mp.PART_NUMBER,\r\n"
			+ "    mp.TOTAL,\r\n"
			+ "    mp.NET_FULFILMENT,\r\n"
			+ "    mp.GROSS_REQ,\r\n"
			+ "    mp.NET_REQ,\r\n"
			+ "    mp.REQ_AHM_OEM,\r\n"
			+ "    mp.REQ_AHM_REM,\r\n"
			+ "    mp.REQ_FDR,\r\n"
			+ "    mp.DIFFERENCE_OFS,\r\n"
			+ "    pat.PATTERN_NAME,       \r\n"
			+ "    sz.DESCRIPTION          \r\n"
			+ "FROM \r\n"
			+ "    SRI_IMPP_D_MONTHLYPLAN_CURING mp\r\n"
			+ "JOIN SRI_IMPP_M_PRODUCT p ON mp.PART_NUMBER = p.PART_NUMBER\r\n"
			+ "LEFT JOIN SRI_IMPP_M_PATTERN pat ON p.PATTERN_ID = pat.PATTERN_ID\r\n"
			+ "LEFT JOIN SRI_IMPP_M_SIZE sz ON p.SIZE_ID = sz.SIZE_ID\r\n"
			+ "ORDER BY \r\n"
			+ "    mp.DETAIL_ID_CURING ASC", nativeQuery = true)
	List<Map<String, Object>> getDetailMP();
	
	@Query(value = "SELECT \r\n"
	        + "    mp.DETAIL_ID_CURING,\r\n"
	        + "    mp.DOC_NUM_CURING,"
	        + "	   mp.PART_NUMBER,\r\n"
	        + "    mp.TOTAL,\r\n"
	        + "    mp.NET_FULFILMENT,\r\n"
	        + "    mp.GROSS_REQ,\r\n"
	        + "    mp.NET_REQ,\r\n"
	        + "    mp.REQ_AHM_OEM,\r\n"
	        + "    mp.REQ_AHM_REM,\r\n"
	        + "    mp.REQ_FDR,\r\n"
	        + "    mp.DIFFERENCE_OFS,\r\n"
	        + "    pat.PATTERN_NAME,       \r\n"
	        + "    sz.DESCRIPTION          \r\n"
	        + "FROM \r\n"
	        + "    SRI_IMPP_D_MONTHLYPLAN_CURING mp\r\n"
	        + "JOIN SRI_IMPP_M_PRODUCT p ON mp.PART_NUMBER = p.PART_NUMBER\r\n"
	        + "LEFT JOIN SRI_IMPP_M_PATTERN pat ON p.PATTERN_ID = pat.PATTERN_ID\r\n"
	        + "LEFT JOIN SRI_IMPP_M_SIZE sz ON p.SIZE_ID = sz.SIZE_ID\r\n"
	        + "WHERE \r\n"
	        + "    mp.DOC_NUM_CURING = :docNum \r\n"
	        + "ORDER BY \r\n"
	        + "    mp.DETAIL_ID_CURING ASC", nativeQuery = true)
	List<Map<String, Object>> getDetailMPById(@Param("docNum") String docNum);
	
	@Query(value = "SELECT dm.*,wd.DESCRIPTION\r\n"
			+ "FROM SRI_IMPP_D_DAILYMONTHLYPLAN_CURING dm\r\n"
			+ "JOIN SRI_IMPP_D_WD_HOURS wd\r\n"
			+ "ON dm.DATE_DAILY_MP = wd.DATE_WD\r\n"
			+ "WHERE dm.DETAIL_ID_CURING = :detailId\r\n"
			+ "ORDER BY dm.DATE_DAILY_MP ASC", nativeQuery = true)
	List<Map<String, Object>> getDetailDailyMP(@Param("detailId")BigDecimal detailId);
	
	@Query(value = "SELECT \r\n"
			+ "    sf.*,\r\n"
			+ "    mo.CHANGE_DATE,\r\n"
			+ "    mo.WORK_CENTER_TEXT as WCT,\r\n"
			+ "    mo.SHIFT\r\n"
			+ "FROM SRI_IMPP_D_SHIFTMONTHLYPLAN_CURING sf JOIN SRI_IMPP_D_DAILYMONTHLYPLAN_CURING dm \r\n"
			+ "ON sf.DETAIL_DAILY_ID_CURING = dm.DETAIL_DAILY_ID_CURING\r\n"
			+ "JOIN SRI_IMPP_D_MONTHLYPLAN_CURING mp \r\n"
			+ "ON dm.DETAIL_ID_CURING = mp.DETAIL_ID_CURING\r\n"
			+ "LEFT JOIN SRI_IMPP_CHANGE_MOULD mo \r\n"
			+ "ON mp.PART_NUMBER = mo.PART_NUMBER AND mo.CHANGE_DATE = sf.ACTUAL_DATE\r\n"
			+ "WHERE sf.DETAIL_DAILY_ID_CURING = :detailDailyId \r\n"
			+ "AND sf.ACTUAL_DATE = :actualDate", nativeQuery = true)
	List<Map<String, Object>> getDetailShiftMP(@Param("actualDate") Date actualDate, @Param("detailDailyId") BigDecimal detailDailyId);
	
	@Query(value = "SELECT DETAIL_ID_CURING,PART_NUMBER,TOTAL,NET_FULFILMENT,GROSS_REQ,NET_REQ,REQ_AHM_OEM,REQ_AHM_REM,REQ_FDR,DIFFERENCE_OFS FROM SRI_IMPP_D_MONTHLYPLAN_CURING where DETAIL_ID_CURING = :detailDailyId\r\n"
	+ " ORDER BY DETAIL_ID_CURING ASC", nativeQuery = true)
	List<Map<String, Object>> getDetailMonthlyPlan(@Param("detailDailyId")BigDecimal detailDailyId);

	@Query(value = "SELECT DESCRIPTION, DATE_WD FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE \r\n"
			+ "EXTRACT(MONTH FROM DATE_WD) = :month\r\n"
			+ "AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
	List<Map<String, Object>> getDescriptionWD(@Param("month")int month, @Param("year")int year);

	
	@Query(value = "SELECT * \r\n"
			+ "FROM SRI_IMPP_D_SHIFTMONTHLYPLAN_CURING \r\n"
			+ "WHERE DETAIL_DAILY_ID_CURING = :detailDailyId\r\n"
			+ "AND EXTRACT(YEAR FROM ACTUAL_DATE) = :year\r\n"
			+ "AND EXTRACT(MONTH FROM ACTUAL_DATE) = :month", nativeQuery = true)
	List<Map<String, Object>>  getDetailShiftMP(@Param("detailDailyId")BigDecimal detailDailyId, @Param("month")String month, @Param("year")String year);
	
	@Query(value = "SELECT \r\n"
			+ "    CURING.DETAIL_DAILY_ID_CURING,\r\n"
			+ "    CURING.DETAIL_ID_CURING,\r\n"
			+ "    CURING.DATE_DAILY_MP,\r\n"
			+ "    CURING.WORK_DAY,\r\n"
			+ "    CURING.TOTAL_PLAN,\r\n"
			+ "    CURING.STATUS,\r\n"
			+ "    WD_HOURS.DESCRIPTION\r\n"
			+ "FROM \r\n"
			+ "    SRI_IMPP_D_DAILYMONTHLYPLAN_CURING CURING\r\n"
			+ "LEFT JOIN \r\n"
			+ "    SRI_IMPP_D_WD_HOURS WD_HOURS\r\n"
			+ "ON \r\n"
			+ "    TO_CHAR(CURING.DATE_DAILY_MP, 'YYYY-MM-DD') = TO_CHAR(WD_HOURS.DATE_WD, 'YYYY-MM-DD')\r\n"
			+ "WHERE \r\n"
			+ "    CURING.DETAIL_ID_CURING = :detailId\r\n"
			+ "ORDER BY \r\n"
			+ "    CURING.DATE_DAILY_MP ASC", nativeQuery = true)
	List<Map<String, Object>> getDetailDailyMonthlyPlan(@Param("detailId")BigDecimal detailId);
	
	
}