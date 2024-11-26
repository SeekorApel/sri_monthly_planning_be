package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.Plant;


public interface ItemCuringRepo extends JpaRepository<ItemCuring, String>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_ITEMCURING WHERE ITEM_CURING = :id", nativeQuery = true)
    Optional<ItemCuring> findById(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_ITEMCURING ORDER BY ITEM_CURING ASC", nativeQuery = true)
    List<ItemCuring> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_ITEMCURING WHERE STATUS = 1", nativeQuery = true)
	List<ItemCuring> findItemCuringActive();
}
