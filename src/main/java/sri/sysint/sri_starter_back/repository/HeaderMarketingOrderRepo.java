package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;

import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.WorkDay; 



public interface HeaderMarketingOrderRepo extends JpaRepository <HeaderMarketingOrder, BigDecimal> {
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_HEADERMARKETINGORDER WHERE HEADER_ID = :id", nativeQuery = true)
    Optional<HeaderMarketingOrder> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_HEADERMARKETINGORDER", nativeQuery = true)
    BigDecimal getNewId();
	

	@Query(value = "SELECT * FROM SRI_IMPP_D_HEADERMARKETINGORDER  WHERE MO_ID = :moId ORDER BY MONTH ASC", nativeQuery = true)
	List<HeaderMarketingOrder> findByMoId(@Param("moId") String moId);

//    @Query(value = "SELECT ROUND(SUM(IWD_SHIFT_1 + IWD_SHIFT_2 + IWD_SHIFT_3) / 3, 2) AS FINAL_WD, "
//            + "ROUND(SUM(IOT_TT_1 + IOT_TT_2 + IOT_TT_3) / 3, 2) AS FINAL_OT_TT, "
//            + "ROUND(SUM(IOT_TL_1 + IOT_TL_2 + IOT_TL_3) / 3, 2) AS FINAL_OT_TL, "
//            + "ROUND((SUM(IWD_SHIFT_1 + IWD_SHIFT_2 + IWD_SHIFT_3) / 3) + (SUM(IOT_TT_1 + IOT_TT_2 + IOT_TT_3) / 3), 2) AS TOTAL_OT_TT, "
//            + "ROUND((SUM(IWD_SHIFT_1 + IWD_SHIFT_2 + IWD_SHIFT_3) / 3) + (SUM(IOT_TL_1 + IOT_TL_2 + IOT_TL_3) / 3), 2) AS TOTAL_OT_TL "
//            + "FROM SRI_IMPP_M_WD "
//            + "WHERE EXTRACT(MONTH FROM DATE_WD) = :month "
//            + "AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
//	Map<String, Object> getMonthlyWorkData(@Param("month") int month, @Param("year") int year);
    
	
    @Query(value = "SELECT \r\n"
    		+ "    ROUND(SUM(CASE WHEN DESCRIPTION = 'WD_NORMAL' THEN (\r\n"
    		+ "        HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "        HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "        HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "        HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "    ) / 24 ELSE 0 END), 2) AS FINAL_WD,\r\n"
    		+ "    \r\n"
    		+ "    ROUND(SUM(CASE WHEN DESCRIPTION = 'OT_TT' THEN (\r\n"
    		+ "        HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "        HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "        HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "        HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "    ) / 24 ELSE 0 END), 2) AS FINAL_OT_TT,\r\n"
    		+ "    \r\n"
    		+ "    ROUND(SUM(CASE WHEN DESCRIPTION = 'OT_TL' THEN (\r\n"
    		+ "        HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "        HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "        HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "        HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "    ) / 24 ELSE 0 END), 2) AS FINAL_OT_TL,\r\n"
    		+ "    \r\n"
    		+ "    ROUND(\r\n"
    		+ "        SUM(CASE WHEN DESCRIPTION = 'WD_NORMAL' THEN (\r\n"
    		+ "            HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "            HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "            HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "            HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "        ) / 24 ELSE 0 END) + \r\n"
    		+ "        SUM(CASE WHEN DESCRIPTION = 'OT_TT' THEN (\r\n"
    		+ "            HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "            HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "            HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "            HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "        ) / 24 ELSE 0 END), 2\r\n"
    		+ "    ) AS TOTAL_OT_TT,\r\n"
    		+ "    \r\n"
    		+ "    ROUND(\r\n"
    		+ "        SUM(CASE WHEN DESCRIPTION = 'WD_NORMAL' THEN (\r\n"
    		+ "            HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "            HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "            HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "            HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "        ) / 24 ELSE 0 END) + \r\n"
    		+ "        SUM(CASE WHEN DESCRIPTION = 'OT_TL' THEN (\r\n"
    		+ "            HOUR_1 + HOUR_2 + HOUR_3 + HOUR_4 + HOUR_5 + HOUR_6 + \r\n"
    		+ "            HOUR_7 + HOUR_8 + HOUR_9 + HOUR_10 + HOUR_11 + HOUR_12 + \r\n"
    		+ "            HOUR_13 + HOUR_14 + HOUR_15 + HOUR_16 + HOUR_17 + HOUR_18 + \r\n"
    		+ "            HOUR_19 + HOUR_20 + HOUR_21 + HOUR_22 + HOUR_23 + HOUR_24\r\n"
    		+ "        ) / 24 ELSE 0 END), 2\r\n"
    		+ "    ) AS TOTAL_OT_TL\r\n"
    		+ "    \r\n"
    		+ "FROM SRI_IMPP_D_WD_HOURS \r\n"
    		+ "WHERE EXTRACT(MONTH FROM DATE_WD) = :month \r\n"
    		+ "  AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
	Map<String, Object> getMonthlyWorkData(@Param("month") int month, @Param("year") int year);
    
    
}
