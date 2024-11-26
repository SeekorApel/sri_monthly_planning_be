package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import sri.sysint.sri_starter_back.model.DetailMarketingOrder;
import sri.sysint.sri_starter_back.model.MarketingOrder;
import sri.sysint.sri_starter_back.model.view.ViewDetailMarketingOrder;

import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

public interface DetailMarketingOrderRepo extends JpaRepository<DetailMarketingOrder, BigDecimal>{
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_MARKETINGORDER WHERE HEADER_ID = :id", nativeQuery = true)
	
    Optional<DetailMarketingOrder> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_MARKETINGORDER WHERE DETAIL_ID = :id", nativeQuery = true)
	DetailMarketingOrder findByDetailId(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT \r\n"
			+ "    t.CATEGORY,\r\n"
			+ "    p.PART_NUMBER,\r\n"
			+ "    p.DESCRIPTION,\r\n"
			+ "    i.MACHINE_TYPE,\r\n"
			+ "    ROUND(i.KAPA_PER_MOULD * (SELECT SETTING_VALUE FROM SRI_IMPP_M_SETTING WHERE SETTING_KEY = 'Capacity' AND STATUS = 1) / 100, 0) AS KAPA_PER_MOULD,\r\n"
			+ "    i.NUMBER_OF_MOULD,\r\n"
			+ "    i.SPARE_MOULD,\r\n"
			+ "    i.MOULD_MONTHLY_PLAN,\r\n"
			+ "    p.QTY_PER_RAK,\r\n"
			+ "    CEIL(p.UPPER_CONSTANT / p.QTY_PER_RAK) * p.QTY_PER_RAK AS MIN_ORDER,\r\n"
			+ "    CASE \r\n"
			+ "        WHEN p.QTY_PER_RAK = 0 THEN 0 \r\n"
			+ "        ELSE FLOOR(\r\n"
			+ "            (ROUND(i.KAPA_PER_MOULD * (SELECT SETTING_VALUE FROM SRI_IMPP_M_SETTING WHERE SETTING_KEY = 'Capacity' AND STATUS = 1) / 100, 0) * i.NUMBER_OF_MOULD * 0.9 * \r\n"
			+ "            (CASE \r\n"
			+ "                WHEN t.PRODUCT_TYPE = 'TT' THEN TO_NUMBER(:totalHKTT1)\r\n"
			+ "                WHEN t.PRODUCT_TYPE = 'TL' THEN TO_NUMBER(:totalHKTL1)\r\n"
			+ "                ELSE 1 -- Nilai default jika tidak TT atau TL\r\n"
			+ "            END)) / p.QTY_PER_RAK\r\n"
			+ "        ) * p.QTY_PER_RAK \r\n"
			+ "    END AS KAPASITAS_MAKSIMUM_1,\r\n"
			+ "    CASE \r\n"
			+ "        WHEN p.QTY_PER_RAK = 0 THEN 0 \r\n"
			+ "        ELSE FLOOR(\r\n"
			+ "            (ROUND(i.KAPA_PER_MOULD * (SELECT SETTING_VALUE FROM SRI_IMPP_M_SETTING WHERE SETTING_KEY = 'Capacity' AND STATUS = 1) / 100, 0) * i.NUMBER_OF_MOULD * 0.9 * \r\n"
			+ "            (CASE \r\n"
			+ "                WHEN t.PRODUCT_TYPE = 'TT' THEN TO_NUMBER(:totalHKTT2)\r\n"
			+ "                WHEN t.PRODUCT_TYPE = 'TL' THEN TO_NUMBER(:totalHKTL2)\r\n"
			+ "                ELSE 1 -- Nilai default jika tidak TT atau TL\r\n"
			+ "            END)) / p.QTY_PER_RAK\r\n"
			+ "        ) * p.QTY_PER_RAK \r\n"
			+ "    END AS KAPASITAS_MAKSIMUM_2,\r\n"
			+ "    CASE \r\n"
			+ "        WHEN p.QTY_PER_RAK = 0 THEN 0 \r\n"
			+ "        ELSE FLOOR(\r\n"
			+ "            (ROUND(i.KAPA_PER_MOULD * (SELECT SETTING_VALUE FROM SRI_IMPP_M_SETTING WHERE SETTING_KEY = 'Capacity' AND STATUS = 1) / 100, 0) * i.NUMBER_OF_MOULD * 0.9 * \r\n"
			+ "            (CASE \r\n"
			+ "                WHEN t.PRODUCT_TYPE = 'TT' THEN TO_NUMBER(:totalHKTT3)\r\n"
			+ "                WHEN t.PRODUCT_TYPE = 'TL' THEN TO_NUMBER(:totalHKTL3)\r\n"
			+ "                ELSE 1 -- Nilai default jika tidak TT atau TL\r\n"
			+ "            END)) / p.QTY_PER_RAK\r\n"
			+ "        ) * p.QTY_PER_RAK \r\n"
			+ "    END AS KAPASITAS_MAKSIMUM_3\r\n"
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
			+ "    t.PRODUCT_MERK = :productMerk AND\r\n"
			+ "    p.STATUS = 1\r\n"
			+ "ORDER BY \r\n"
			+ "    t.CATEGORY", nativeQuery = true)

	List<Map<String, Object>> getDataTable(
	        @Param("totalHKTT1") BigDecimal totalHKTT1,
	        @Param("totalHKTT2") BigDecimal totalHKTT2,
	        @Param("totalHKTT3") BigDecimal totalHKTT3, 
	        @Param("totalHKTL1") BigDecimal totalHKTL1, 
	        @Param("totalHKTL2") BigDecimal totalHKTL2, 
	        @Param("totalHKTL3") BigDecimal totalHKTL3, 
	        @Param("productMerk") String productMerk);

	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_MARKETINGORDER", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_MARKETINGORDER WHERE MO_ID = :moId", nativeQuery = true)
	List<DetailMarketingOrder> findByMoId(@Param("moId") String moId);

	   
	@Query(value = "SELECT t.CATEGORY, p.DESCRIPTION, i.MACHINE_TYPE " +
	        "FROM SRI_IMPP_M_PRODUCT p " +
	        "JOIN SRI_IMPP_M_ITEMCURING i ON p.ITEM_CURING = i.ITEM_CURING " +
	        "JOIN SRI_IMPP_M_PRODUCTTYPE t ON p.PRODUCT_TYPE_ID = t.PRODUCT_TYPE_ID " +
	        "WHERE p.PART_NUMBER = :partNumber", nativeQuery = true)
	Map<String, Object> findProductDetails(@Param("partNumber") BigDecimal partNumber);



}
