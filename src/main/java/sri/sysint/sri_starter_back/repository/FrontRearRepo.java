package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.transaction.annotation.Transactional;

import sri.sysint.sri_starter_back.model.FrontRear;

public interface FrontRearRepo extends JpaRepository<FrontRear, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_FRONT_REAR WHERE ID_FRONT_REAR = :id", nativeQuery = true)
    Optional<FrontRear> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_FRONT_REAR WHERE ID_FRONT_REAR = :id", nativeQuery = true)
	List<FrontRear> findListById(@Param("id") BigDecimal id);
	
	@Modifying
    @Transactional 
    @Query(value = "DELETE FROM SRI_IMPP_M_FRONT_REAR", nativeQuery = true)
    void deleteAllFr();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_FRONT_REAR ORDER BY ID_FRONT_REAR ASC", nativeQuery = true)
    List<FrontRear> getDataOrderIdFrontRear();
	
    @Query(value = "SELECT COALESCE(MAX(ID_FRONT_REAR), 0) + 1 FROM SRI_IMPP_M_FRONT_REAR", nativeQuery = true)
    BigDecimal getNextId();
	
    @Query(value = "SELECT * FROM SRI_IMPP_M_FRONT_REAR WHERE STATUS = :status", nativeQuery = true)
    List<FrontRear> findByStatus(@Param("status") BigDecimal status);
    
        @Query(value = "SELECT fr.ID_FRONT_REAR, fr.DETAIL_ID_MO, fr.STATUS, fr.CREATION_DATE, fr.CREATED_BY, fr.LAST_UPDATE_DATE, fr.LAST_UPDATED_BY,\r\n"
        		+ "                       mo.DETAIL_ID, mo.MO_ID, mo.CATEGORY, mo.PART_NUMBER, mo.DESCRIPTION, mo.MACHINE_TYPE, mo.CAPACITY, mo.QTY_PER_MOULD,\r\n"
        		+ "                       mo.QTY_PER_RAK, mo.MIN_ORDER, mo.MAX_CAP_MONTH_0, mo.MAX_CAP_MONTH_1, mo.MAX_CAP_MONTH_2, mo.INITIAL_STOCK,\r\n"
        		+ "                       mo.SF_MONTH_0, mo.SF_MONTH_1, mo.SF_MONTH_2, mo.MO_MONTH_0, mo.MO_MONTH_1, mo.MO_MONTH_2, mo.PPD, mo.CAV,\r\n"
        		+ "                       mo.STATUS AS MO_STATUS, mo.CREATION_DATE AS MO_CREATION_DATE, mo.CREATED_BY AS MO_CREATED_BY, mo.LAST_UPDATE_DATE AS MO_LAST_UPDATE_DATE,\r\n"
        		+ "                       mo.LAST_UPDATED_BY AS MO_LAST_UPDATED_BY\r\n"
        		+ "                FROM SRI_IMPP_M_FRONT_REAR fr\r\n"
        		+ "                JOIN SRI_IMPP_D_MARKETINGORDER mo\r\n"
        		+ "                  ON fr.DETAIL_ID_MO = mo.DETAIL_ID\r\n"
        		+ "                WHERE fr.ID_FRONT_REAR = :idFrontRear", nativeQuery = true)
            List<Map<String, Object>> findMarketingOrderByFrontRearId(@Param("idFrontRear") BigDecimal idFrontRear);

            @Query(value = "SELECT fr.ID_FRONT_REAR, fr.DETAIL_ID_MO, fr.STATUS, fr.CREATION_DATE, fr.CREATED_BY, fr.LAST_UPDATE_DATE, fr.LAST_UPDATED_BY,\r\n"
            		+ "                       mo.DETAIL_ID, mo.MO_ID, mo.CATEGORY, mo.PART_NUMBER, mo.DESCRIPTION, mo.MACHINE_TYPE, mo.CAPACITY, mo.QTY_PER_MOULD,\r\n"
            		+ "                       mo.QTY_PER_RAK, mo.MIN_ORDER, mo.MAX_CAP_MONTH_0, mo.MAX_CAP_MONTH_1, mo.MAX_CAP_MONTH_2, mo.INITIAL_STOCK,\r\n"
            		+ "                       mo.SF_MONTH_0, mo.SF_MONTH_1, mo.SF_MONTH_2, mo.MO_MONTH_0, mo.MO_MONTH_1, mo.MO_MONTH_2, mo.PPD, mo.CAV,\r\n"
            		+ "                       mo.STATUS AS MO_STATUS, mo.CREATION_DATE AS MO_CREATION_DATE, mo.CREATED_BY AS MO_CREATED_BY, mo.LAST_UPDATE_DATE AS MO_LAST_UPDATE_DATE,\r\n"
            		+ "                       mo.LAST_UPDATED_BY AS MO_LAST_UPDATED_BY\r\n"
            		+ "                FROM SRI_IMPP_M_FRONT_REAR fr\r\n"
            		+ "                JOIN SRI_IMPP_D_MARKETINGORDER mo\r\n"
            		+ "                  ON fr.DETAIL_ID_MO = mo.DETAIL_ID", nativeQuery = true)
            List<Map<String, Object>> findAllDetailFrMarketingOrders();
            
            @Modifying
            @Transactional
            @Query(value = "INSERT INTO SRI_IMPP_M_FRONT_REAR (ID_FRONT_REAR, DETAIL_ID_MO, STATUS, CREATION_DATE, LAST_UPDATE_DATE) " +
                           "VALUES (:id, :detailIdMo, :status, :creationDate, :lastUpdateDate)", nativeQuery = true)
            int insertNew(
                @Param("id") BigDecimal id,
                @Param("detailIdMo") BigDecimal detailIdMo,
                @Param("status") BigDecimal status,
                @Param("creationDate") Date creationDate,
                @Param("lastUpdateDate") Date lastUpdateDate
            );
}
