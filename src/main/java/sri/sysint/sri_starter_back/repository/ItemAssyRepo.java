package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.ItemAssy;
import sri.sysint.sri_starter_back.model.Plant;


public interface ItemAssyRepo extends JpaRepository<ItemAssy, String> {
    @Query(value = "SELECT * FROM SRI_IMPP_M_ITEMASSY WHERE ITEM_ASSY = :id", nativeQuery = true)
    Optional<ItemAssy> findById(@Param("id") String id);
    
	@Query(value = "SELECT * FROM SRI_IMPP_M_ITEMASSY ORDER BY ITEM_ASSY ASC", nativeQuery = true)
    List<ItemAssy> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_ITEMASSY WHERE STATUS = 1", nativeQuery = true)
	List<ItemAssy> findItemAssyActive();
}

