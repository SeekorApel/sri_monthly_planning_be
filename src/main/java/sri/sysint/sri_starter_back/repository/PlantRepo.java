package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Plant;

public interface PlantRepo extends JpaRepository<Plant, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_PLANT WHERE PLANT_ID = :id", nativeQuery = true)
    Optional<Plant> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PLANT ORDER BY PLANT_ID ASC", nativeQuery = true)
    List<Plant> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_PLANT", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PLANT WHERE STATUS = 1", nativeQuery = true)
	List<Plant> findPlantActive();

	@Query(value = "SELECT * FROM SRI_IMPP_M_PLANT WHERE PLANT_NAME = :name", nativeQuery = true)
	Optional<Plant> findByName(@Param("name") String name);

}
