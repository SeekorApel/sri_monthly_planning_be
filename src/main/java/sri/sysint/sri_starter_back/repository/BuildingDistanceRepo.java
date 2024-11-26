package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.Plant;

public interface BuildingDistanceRepo extends JpaRepository<BuildingDistance, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_D_BUILDINGDISTANCE WHERE ID_B_DISTANCE = :id", nativeQuery = true)
    Optional<BuildingDistance> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_BUILDINGDISTANCE ORDER BY ID_B_DISTANCE ASC", nativeQuery = true)
    List<BuildingDistance> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_BUILDINGDISTANCE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_BUILDINGDISTANCE WHERE STATUS = 1", nativeQuery = true)
	List<BuildingDistance> findBuildingDistanceActive();
}
