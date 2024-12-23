package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.DWorkDayHoursSpesific;

public interface DWorkDayHoursSpecificRepo extends JpaRepository<DWorkDayHoursSpesific, Date> {
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE DETAIL_WD_HOURS_SPECIFIC_ID = :id", nativeQuery = true)
    Optional<DWorkDayHoursSpesific> findById(@Param("id") BigDecimal id);

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE TRUNC(DATE_WD) = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	Optional<DWorkDayHoursSpesific> findDWdHoursByDate(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE TRUNC(DATE_WD) = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	List<DWorkDayHoursSpesific> findDWdHoursListByDate(@Param("id") String id);


    @Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC ORDER BY DATE_WD ASC", nativeQuery = true)
    List<DWorkDayHoursSpesific> getDataOrderByDateDWd();
    
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE STATUS = 1", nativeQuery = true)
	List<DWorkDayHoursSpesific> findDWorkDayHoursActive();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_WD_HOURS_SPECIFIC", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE TO_DATE(DATE_WD, 'DD-MM-YYYY') = TO_DATE(:date, 'DD-MM-YYYY') AND DESCRIPTION = :description", nativeQuery = true)
	Optional<DWorkDayHoursSpesific> findDWdHoursByDateAndDescription(@Param("date") String date, @Param("description") String description);

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC", nativeQuery = true)
	List<DWorkDayHoursSpesific> fetchAllManual();

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE EXTRACT(MONTH FROM DATE_WD) = :month AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
	List<DWorkDayHoursSpesific> findDWdHoursByMonthAndYear(@Param("month") int month, @Param("year") int year);
	
	@Query(value = "SELECT "
			+ "    TO_CHAR(DATE_WD, 'DD-MM-YYYY') AS DATE_WD, "
			+ "    ROUND(SHIFT1_TOTAL_TIME / 60, 3) AS SHIFT1_HOUR, "
			+ "    ROUND(SHIFT2_TOTAL_TIME / 60, 3) AS SHIFT2_HOUR, "
			+ "    ROUND(SHIFT3_TOTAL_TIME / 60, 3) AS SHIFT3_HOUR, "
			+ "    ROUND(SHIFT1_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT1_CAPACITY, "
			+ "    ROUND(SHIFT2_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT2_CAPACITY, "
			+ "    ROUND(SHIFT3_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT3_CAPACITY, "
			+ "    ROUND((SHIFT1_TOTAL_TIME - 120) / 60, 3) AS SHIFT1_HOUR_CHANGE_MOULD, "
			+ "    ROUND((SHIFT2_TOTAL_TIME - 120) / 60, 3) AS SHIFT2_HOUR_CHANGE_MOULD, "
			+ "    ROUND((SHIFT3_TOTAL_TIME - 120) / 60, 3) AS SHIFT3_HOUR_CHANGE_MOULD, "
			+ "    ROUND((SHIFT1_TOTAL_TIME - 120) * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT1_CAPACITY_CHANGE_MOULD, "
			+ "    ROUND((SHIFT2_TOTAL_TIME - 120) * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT2_CAPACITY_CHANGE_MOULD, "
			+ "    ROUND((SHIFT3_TOTAL_TIME - 120) * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT3_CAPACITY_CHANGE_MOULD "
			+ "FROM "
			+ "    SRI_IMPP_D_WD_HOURS_SPECIFIC "
			+ "WHERE "
			+ "    DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
			+ "    AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
			+ "    AND ( "
			+ "        DESCRIPTION = 'WD_NORMAL' "
			+ "        OR ( "
			+ "            (:itemCuring LIKE 'CN-%' AND DESCRIPTION = 'OT_TL') "
			+ "            OR (:itemCuring LIKE 'CM-%' AND DESCRIPTION = 'OT_TT') "
			+ "        ) "
			+ "    ) "
			+ "    AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)",  nativeQuery = true)
	List<Map<String, Object>> getCuringCapacity(@Param("itemCuring") String itemCuring, @Param("operationShortText") String operationShortText, @Param("cavity") int cavity, @Param("month") int month, @Param("year") int year);
	
	@Query(value = "SELECT "
			+ "			TO_CHAR(DATE_WD, 'DD-MM-YYYY') AS DATE_WD, "
			+ "      CASE WHEN DATE_WD = ( "
			+ "            SELECT MIN(DATE_WD) "
			+ "            FROM SRI_IMPP_D_WD_HOURS_SPECIFIC "
			+ "            WHERE "
			+ "            DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
			+ "            AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
			+ "            AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)) "
			+ "      THEN ROUND((SHIFT1_TOTAL_TIME - 120) / 60, 3) "
			+ "      ELSE ROUND(SHIFT1_TOTAL_TIME / 60, 3) "
			+ "      END AS SHIFT1_HOUR , "
			+ "			ROUND(SHIFT2_TOTAL_TIME / 60, 3) AS SHIFT2_HOUR, "
			+ "			ROUND(SHIFT3_TOTAL_TIME / 60, 3) AS SHIFT3_HOUR, "
			+ "      CASE WHEN DATE_WD = ( "
			+ "            SELECT MIN(DATE_WD) "
			+ "            FROM SRI_IMPP_D_WD_HOURS_SPECIFIC "
			+ "            WHERE "
			+ "            DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
			+ "            AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
			+ "            AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)) "
			+ "      THEN ROUND((SHIFT1_TOTAL_TIME - 120) * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity"
			+ "      ELSE ROUND(SHIFT1_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity "
			+ "      END AS SHIFT1_CAPACITY , "
			+ "			ROUND(SHIFT2_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT2_CAPACITY, "
			+ "			ROUND(SHIFT3_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT3_CAPACITY "
			+ "			FROM "
			+ "			SRI_IMPP_D_WD_HOURS_SPECIFIC "
			+ "			WHERE "
			+ "			DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
			+ "			AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
			+ "			AND ( "
			+ "			    DESCRIPTION = 'WD_NORMAL' "
			+ "			    OR ( "
			+ "			        (:itemCuring LIKE 'CN-%' AND DESCRIPTION = 'OT_TL') "
			+ "			        OR (:itemCuring LIKE 'CM-%' AND DESCRIPTION = 'OT_TT') "
			+ "			    ) "
			+ "			) "
			+ "			AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)",  nativeQuery = true)
	List<Map<String, Object>> getCuringCapacityChangeMouldFirstDate(@Param("itemCuring") String itemCuring, @Param("operationShortText") String operationShortText, @Param("cavity") int cavity, @Param("month") int month, @Param("year") int year);
	
	@Query(value = "SELECT "
	        + "    TO_CHAR(DATE_WD, 'DD-MM-YYYY') AS DATE_WD, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND :shift = 3 THEN ROUND(SHIFT1_TOTAL_TIME / 60, 3) "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND(SHIFT1_TOTAL_TIME / 60, 3) "
	        + "        ELSE 0 "
	        + "    END AS SHIFT1_HOUR, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift = 1 OR :shift = 3) THEN ROUND(SHIFT2_TOTAL_TIME / 60, 3) "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND(SHIFT2_TOTAL_TIME / 60, 3) "
	        + "        ELSE 0 "
	        + "    END AS SHIFT2_HOUR, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift IN (1, 2, 3)) THEN ROUND(SHIFT3_TOTAL_TIME / 60, 3) "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND(SHIFT3_TOTAL_TIME / 60, 3) "
	        + "        ELSE 0 "
	        + "    END AS SHIFT3_HOUR, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND :shift = 3 THEN ROUND(SHIFT1_TOTAL_TIME * 60 / CT_SEC1) * :cavity "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND(SHIFT1_TOTAL_TIME * 60 / CT_SEC1) * :cavity "
	        + "        ELSE 0 "
	        + "    END AS SHIFT1_CAPACITY, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift = 1 OR :shift = 3) THEN ROUND(SHIFT2_TOTAL_TIME * 60 / CT_SEC1) * :cavity "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND(SHIFT2_TOTAL_TIME * 60 / CT_SEC1) * :cavity "
	        + "        ELSE 0 "
	        + "    END AS SHIFT2_CAPACITY, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift IN (1, 2, 3)) THEN ROUND(SHIFT3_TOTAL_TIME * 60 / CT_SEC1) * :cavity "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND(SHIFT3_TOTAL_TIME * 60 / CT_SEC1) * :cavity "
	        + "        ELSE 0 "
	        + "    END AS SHIFT3_CAPACITY, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND :shift = 3 THEN ROUND((SHIFT1_TOTAL_TIME - 120) / 60, 3) "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND((SHIFT1_TOTAL_TIME - 120) / 60, 3) "
	        + "        ELSE 0 "
	        + "    END AS SHIFT1_HOUR_CHANGE_MOULD, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift = 1 OR :shift = 3) THEN ROUND((SHIFT2_TOTAL_TIME - 120) / 60, 3) "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND((SHIFT2_TOTAL_TIME - 120) / 60, 3) "
	        + "        ELSE 0 "
	        + "    END AS SHIFT2_HOUR_CHANGE_MOULD, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift IN (1, 2, 3)) THEN ROUND((SHIFT3_TOTAL_TIME - 120) / 60, 3) "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND((SHIFT3_TOTAL_TIME - 120) / 60, 3) "
	        + "        ELSE 0 "
	        + "    END AS SHIFT3_HOUR_CHANGE_MOULD, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND :shift = 3 THEN ROUND((SHIFT1_TOTAL_TIME - 120) * 60 / CT_SEC1) * :cavity "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND((SHIFT1_TOTAL_TIME - 120) * 60 / CT_SEC1) * :cavity "
	        + "        ELSE 0 "
	        + "    END AS SHIFT1_CAPACITY_CHANGE_MOULD, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift = 1 OR :shift = 3) THEN ROUND((SHIFT2_TOTAL_TIME - 120) * 60 / CT_SEC1) * :cavity "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND((SHIFT2_TOTAL_TIME - 120) * 60 / CT_SEC1) * :cavity "
	        + "        ELSE 0 "
	        + "    END AS SHIFT2_CAPACITY_CHANGE_MOULD, "
	        + "    CASE "
	        + "        WHEN DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND (:shift IN (1, 2, 3)) THEN ROUND((SHIFT3_TOTAL_TIME - 120) * 60 / CT_SEC1) * :cavity "
	        + "        WHEN DATE_WD != TO_DATE(:date, 'DD-MM-YYYY') THEN ROUND((SHIFT3_TOTAL_TIME - 120) * 60 / CT_SEC1) * :cavity "
	        + "        ELSE 0 "
	        + "    END AS SHIFT3_CAPACITY_CHANGE_MOULD "
	        + "FROM "
	        + "    SRI_IMPP_D_WD_HOURS_SPECIFIC WD "
	        + "    LEFT JOIN ( "
	        + "        SELECT WIP, OPERATION_SHORT_TEXT, CT_SEC1 "
	        + "        FROM SRI_IMPP_M_CT_CURING "
	        + "        WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText "
	        + "    ) CURING ON 1 = 1 "
	        + "WHERE "
	        + "    DATE_WD BETWEEN "
	        + "        CASE WHEN :shift = 3 THEN TO_DATE(:date, 'DD-MM-YYYY') + 1 ELSE TO_DATE(:date, 'DD-MM-YYYY') END "
	        + "    AND LAST_DAY(TO_DATE(:date, 'DD-MM-YYYY')) "
	        + "    AND ( "
	        + "        DESCRIPTION = 'WD_NORMAL' "
	        + "        OR ( "
	        + "            (:itemCuring LIKE 'CN-%' AND DESCRIPTION = 'OT_TL') "
	        + "            OR (:itemCuring LIKE 'CM-%' AND DESCRIPTION = 'OT_TT') "
	        + "        ) "
	        + "    ) "
	        + "    AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)", nativeQuery = true)
	List<Map<String, Object>> getCuringCapacityMidMonth(@Param("itemCuring") String itemCuring, 
	                                                @Param("operationShortText") String operationShortText, 
	                                                @Param("cavity") int cavity, 
	                                                @Param("date") String date, 
	                                                @Param("shift") int shift);
	
	@Query(value = "SELECT "
			+ "		TO_CHAR(DATE_WD, 'DD-MM-YYYY') AS DATE_WD,"
            + "     ROUND(SHIFT1_TOTAL_TIME / 60, 3) AS SHIFT1_HOUR, "
            + "     ROUND(SHIFT2_TOTAL_TIME / 60, 3) AS SHIFT2_HOUR, "
            + "     ROUND(SHIFT3_TOTAL_TIME / 60, 3) AS SHIFT3_HOUR, "
            + "     ROUND(SHIFT1_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT1_CAPACITY, "
            + "     ROUND(SHIFT2_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT2_CAPACITY, "
            + "     ROUND(SHIFT3_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT3_CAPACITY "
            + " FROM "
            + "     SRI_IMPP_D_WD_HOURS_SPECIFIC "
            + " WHERE "
            + "     DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
            + "     AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
            + "     AND DESCRIPTION = 'WD_NORMAL' "
            + "     AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)",  nativeQuery = true)
	List<Map<String, Object>> getCuringCapacityWdNormal(@Param("itemCuring") String itemCuring, @Param("operationShortText") String operationShortText, @Param("cavity") int cavity, @Param("month") int month, @Param("year") int year);
	
	@Query(value = "SELECT "
			+ "		TO_CHAR(DATE_WD, 'DD-MM-YYYY') AS DATE_WD,"
            + "     ROUND(SHIFT1_TOTAL_TIME / 60, 3) AS SHIFT1_HOUR, "
            + "     ROUND(SHIFT2_TOTAL_TIME / 60, 3) AS SHIFT2_HOUR, "
            + "     ROUND(SHIFT3_TOTAL_TIME / 60, 3) AS SHIFT3_HOUR, "
            + "     ROUND(SHIFT1_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT1_CAPACITY, "
            + "     ROUND(SHIFT2_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT2_CAPACITY, "
            + "     ROUND(SHIFT3_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT3_CAPACITY "
            + " FROM "
            + "     SRI_IMPP_D_WD_HOURS_SPECIFIC "
            + " WHERE "
            + "     DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
            + "     AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
            + "     AND DESCRIPTION = 'OT_TL' "
            + "     AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)",  nativeQuery = true)
	List<Map<String, Object>> getCuringCapacityOtTl(@Param("itemCuring") String itemCuring, @Param("operationShortText") String operationShortText, @Param("cavity") int cavity, @Param("month") int month, @Param("year") int year);
	
	@Query(value = "SELECT "
			+ "		TO_CHAR(DATE_WD, 'DD-MM-YYYY') AS DATE_WD,"
            + "     ROUND(SHIFT1_TOTAL_TIME / 60, 3) AS SHIFT1_HOUR, "
            + "     ROUND(SHIFT2_TOTAL_TIME / 60, 3) AS SHIFT2_HOUR, "
            + "     ROUND(SHIFT3_TOTAL_TIME / 60, 3) AS SHIFT3_HOUR, "
            + "     ROUND(SHIFT1_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT1_CAPACITY, "
            + "     ROUND(SHIFT2_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT2_CAPACITY, "
            + "     ROUND(SHIFT3_TOTAL_TIME * 60 / (SELECT CT_SEC1 FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring AND OPERATION_SHORT_TEXT = :operationShortText)) * :cavity AS SHIFT3_CAPACITY "
            + " FROM "
            + "     SRI_IMPP_D_WD_HOURS_SPECIFIC "
            + " WHERE "
            + "     DATE_WD BETWEEN TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY') "
            + "     AND LAST_DAY(TO_DATE('01-' || :month || '-' || :year, 'DD-MM-YYYY')) "
            + "     AND DESCRIPTION = 'OT_TT' "
            + "     AND (SHIFT1_TOTAL_TIME > 0 OR SHIFT2_TOTAL_TIME > 0 OR SHIFT3_TOTAL_TIME > 0)",  nativeQuery = true)
	List<Map<String, Object>> getCuringCapacityOtTt(@Param("itemCuring") String itemCuring, @Param("operationShortText") String operationShortText, @Param("cavity") int cavity, @Param("month") int month, @Param("year") int year);

}
