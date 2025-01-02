package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;

import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder; 



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
    
	
	@Query(value = "SELECT "
    		+ "    		  ROUND(SUM(CASE WHEN DESCRIPTION = 'WD_NORMAL' THEN ( "
    		+ "    		      SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "    		  ) / 60 / 24 ELSE 0 END), 2) AS FINAL_WD, "
    		+ "    		  "
    		+ "    		  ROUND(SUM(CASE WHEN DESCRIPTION = 'OT_TT' THEN ( "
    		+ "    		      SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "    		  ) / 60 / 24 ELSE 0 END), 2) AS FINAL_OT_TT, "
    		+ "    		  "
    		+ "    		  ROUND(SUM(CASE WHEN DESCRIPTION = 'OT_TL' THEN ( "
    		+ "    		      SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "    		  ) / 60 / 24 ELSE 0 END), 2) AS FINAL_OT_TL, "
    		+ "    		  "
    		+ "    		  ROUND( "
    		+ "    		      SUM(CASE WHEN DESCRIPTION = 'WD_NORMAL' THEN ( "
    		+ "    		          SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "              ) / 60 / 24 ELSE 0 END) + "
    		+ "    		      SUM(CASE WHEN DESCRIPTION = 'OT_TT' THEN ( "
    		+ "    		          SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "              ) / 60 / 24 ELSE 0 END), 2 "
    		+ "    		  ) AS TOTAL_OT_TT, "
    		+ "    		  "
    		+ "    		  ROUND( "
    		+ "    		      SUM(CASE WHEN DESCRIPTION = 'WD_NORMAL' THEN ( "
    		+ "    		          SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "              ) / 60 / 24 ELSE 0 END) + "
    		+ "    		      SUM(CASE WHEN DESCRIPTION = 'OT_TL' THEN ( "
    		+ "    		          SHIFT1_TOTAL_TIME + SHIFT2_TOTAL_TIME + SHIFT3_TOTAL_TIME "
    		+ "              ) / 60 / 24 ELSE 0 END), 2 "
    		+ "    		  ) AS TOTAL_OT_TL "
    		+ "    		  "
    		+ "    		FROM SRI_IMPP_D_WD_HOURS_SPECIFIC "
    		+ "    		WHERE EXTRACT(MONTH FROM DATE_WD) = :month "
    		+ "    		AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
	Map<String, Object> getMonthlyWorkData(@Param("month") int month, @Param("year") int year);
	
	//HeaderMarketingOrderRepo
	@Query(value = "SELECT * FROM SRI_IMPP_D_HEADERMARKETINGORDER WHERE MO_ID = :moId1 OR MO_ID = :moId2", nativeQuery = true)
	List<HeaderMarketingOrder> findByTwoMoId(@Param("moId1") String moId1, @Param("moId2") String moId2);
   

}
