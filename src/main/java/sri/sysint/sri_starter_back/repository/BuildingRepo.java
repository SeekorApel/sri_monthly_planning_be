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

public interface BuildingRepo extends JpaRepository<Building, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_BUILDING WHERE BUILDING_ID = :id", nativeQuery = true)
    Optional<Building> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_BUILDING ORDER BY BUILDING_ID ASC", nativeQuery = true)
    List<Building> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_BUILDING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_BUILDING WHERE STATUS = 1", nativeQuery = true)
	List<Building> findBuildingActive();
	
    @Query(value = "SELECT * FROM SRI_IMPP_M_BUILDING WHERE BUILDING_NAME = :name", nativeQuery = true)
    Optional<Building> findByName(@Param("name") String name);
}
