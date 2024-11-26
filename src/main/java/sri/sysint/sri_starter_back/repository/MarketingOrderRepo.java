package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MarketingOrder;

public interface MarketingOrderRepo extends JpaRepository<MarketingOrder, String>{
	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MARKETINGORDER WHERE MO_ID = :id", nativeQuery = true)
	Optional<MarketingOrder> findById(@Param("id") String id);
	
//	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_T_MARKETINGORDER", nativeQuery = true)
//	String getNewId();
	
	@Query(value = "SELECT * FROM (SELECT * FROM SRI_IMPP_T_MARKETINGORDER ORDER BY MO_ID DESC) WHERE ROWNUM = 1", nativeQuery = true)
	MarketingOrder findLastMOId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MARKETINGORDER WHERE MO_ID = :moId", nativeQuery = true)
	MarketingOrder findByMoId(@Param("moId") String moId);

	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_T_MARKETINGORDER WHERE MONTH_0 = :month0 AND MONTH_1 = :month1 AND MONTH_2 = :month2 AND TYPE = :type", nativeQuery = true)
	BigDecimal getNewRevisionPpc(@Param("month0") Date month0, @Param("month1") Date month1, @Param("month2") Date month2, @Param("type") String type);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_T_MARKETINGORDER WHERE MONTH_0 = :month0 AND MONTH_1 = :month1 AND MONTH_2 = :month2 AND TYPE = :type", nativeQuery = true)
	BigDecimal getNewRevisionMarketing(@Param("month0") Date month0, @Param("month1") Date month1, @Param("month2") Date month2, @Param("type") String type);
	
	//Add dicky
	@Query(value = "SELECT MO_ID FROM SRI_IMPP_T_MARKETINGORDER ORDER BY MO_ID DESC FETCH FIRST 1 ROWS ONLY", nativeQuery = true)
	String getLastIdMo();
	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MARKETINGORDER ORDER BY MO_ID DESC", nativeQuery = true)
    List<MarketingOrder> findAllMOByIdDesc();
	
	@Query(value = "SELECT *\r\n"
			+ "FROM SRI_IMPP_T_MARKETINGORDER t1\r\n"
			+ "WHERE MO_ID = (\r\n"
			+ "    SELECT MAX(t2.MO_ID)\r\n"
			+ "    FROM SRI_IMPP_T_MARKETINGORDER t2\r\n"
			+ "    WHERE t2.TYPE = t1.TYPE\r\n"
			+ "      AND TO_CHAR(t2.MONTH_0, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_0, 'DD-MM-YYYY')\r\n"
			+ "      AND TO_CHAR(t2.MONTH_1, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_1, 'DD-MM-YYYY')\r\n"
			+ "      AND TO_CHAR(t2.MONTH_2, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_2, 'DD-MM-YYYY')\r\n"
			+ ")\r\n"
			+ "ORDER BY MO_ID DESC", 
	        nativeQuery = true)
	List<MarketingOrder> findLatestMarketingOrders();
	
	@Query(value = "SELECT *\r\n"
			+ "			 FROM SRI_IMPP_T_MARKETINGORDER t1\r\n"
			+ "			 WHERE MO_ID = (\r\n"
			+ "			     SELECT MAX(t2.MO_ID)\r\n"
			+ "			     FROM SRI_IMPP_T_MARKETINGORDER t2\r\n"
			+ "			     WHERE t2.TYPE = t1.TYPE\r\n"
			+ "			       AND TO_CHAR(t2.MONTH_0, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_0, 'DD-MM-YYYY')\r\n"
			+ "			       AND TO_CHAR(t2.MONTH_1, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_1, 'DD-MM-YYYY')\r\n"
			+ "			       AND TO_CHAR(t2.MONTH_2, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_2, 'DD-MM-YYYY')\r\n"
			+ "             AND t1.TYPE = 'FED'\r\n"
			+ "			 )\r\n"
			+ "			 ORDER BY MO_ID DESC", 
	        nativeQuery = true)
	List<MarketingOrder> findLatestMarketingOrderFED();
	
	@Query(value = "SELECT *\r\n"
			+ "			 FROM SRI_IMPP_T_MARKETINGORDER t1\r\n"
			+ "			 WHERE MO_ID = (\r\n"
			+ "			     SELECT MAX(t2.MO_ID)\r\n"
			+ "			     FROM SRI_IMPP_T_MARKETINGORDER t2\r\n"
			+ "			     WHERE t2.TYPE = t1.TYPE\r\n"
			+ "			       AND TO_CHAR(t2.MONTH_0, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_0, 'DD-MM-YYYY')\r\n"
			+ "			       AND TO_CHAR(t2.MONTH_1, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_1, 'DD-MM-YYYY')\r\n"
			+ "			       AND TO_CHAR(t2.MONTH_2, 'DD-MM-YYYY') = TO_CHAR(t1.MONTH_2, 'DD-MM-YYYY')\r\n"
			+ "             AND t1.TYPE = 'FDR'\r\n"
			+ "			 )\r\n"
			+ "			 ORDER BY MO_ID DESC", 
	        nativeQuery = true)
	List<MarketingOrder> findLatestMarketingOrderFDR();
	
	
	@Query(value = "SELECT * "
	        + "FROM SRI_IMPP_T_MARKETINGORDER "
	        + "WHERE TO_DATE(MONTH_0, 'DD-MM-YYYY') = TO_DATE(:month0, 'DD-MM-YYYY') "
	        + "AND TO_DATE(MONTH_1, 'DD-MM-YYYY') = TO_DATE(:month1, 'DD-MM-YYYY') "
	        + "AND TO_DATE(MONTH_2, 'DD-MM-YYYY') = TO_DATE(:month2, 'DD-MM-YYYY') "
	        + "AND TYPE = :type", 
	        nativeQuery = true)
	
	List<MarketingOrder> findtMarketingOrders(@Param("month0") String month0, 
	                                          @Param("month1") String month1, 
	                                          @Param("month2") String month2, 
	                                          @Param("type") String type);

	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MARKETINGORDER \r\n"
			+ "	    WHERE EXTRACT(MONTH FROM MONTH_0) = :month1  \r\n"
			+ "     AND EXTRACT(YEAR FROM MONTH_0) = :year1\r\n"
			+ "     AND EXTRACT(MONTH FROM MONTH_1) = :month2  \r\n"
			+ "	    AND EXTRACT(YEAR FROM MONTH_1) = :year2\r\n"
			+ "     AND EXTRACT(MONTH FROM MONTH_2) = :month3  \r\n"
			+ "	    AND EXTRACT(YEAR FROM MONTH_2) = :year3\r\n"
			+ "	    AND TYPE = :type", 
	        nativeQuery = true)
	
	List<MarketingOrder> checktMarketingOrders(@Param("month1") String month1, 
	                                          @Param("month2") String month2, 
	                                          @Param("month3") String month3, 
	                                          @Param("year1") String year1, 
	                                          @Param("year2") String year2, 
	                                          @Param("year3") String year3, 
	                                          @Param("type") String type);
	
	@Query(value = "SELECT * FROM SRI_IMPP_T_MARKETINGORDER \r\n"
			+ "WHERE TO_CHAR(MONTH_0, 'MM-YYYY') = :month0 \r\n"
			+ "AND TO_CHAR(MONTH_1, 'MM-YYYY') = :month1 \r\n"
			+ "AND TO_CHAR(MONTH_2, 'MM-YYYY') = :month2", nativeQuery = true)
	
	List<MarketingOrder> findByMonth(@Param("month0") String month0, @Param("month1") String month1, @Param("month2") String month2);


}
