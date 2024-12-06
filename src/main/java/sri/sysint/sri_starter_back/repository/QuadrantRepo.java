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
import sri.sysint.sri_starter_back.model.Quadrant;

public interface QuadrantRepo extends JpaRepository<Quadrant, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_QUADRANT WHERE QUADRANT_ID = :id", nativeQuery = true)
    Optional<Quadrant> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_QUADRANT ORDER BY QUADRANT_ID ASC", nativeQuery = true)
    List<Quadrant> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_QUADRANT", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_QUADRANT WHERE STATUS = 1", nativeQuery = true)
	List<Quadrant> findQuadrantActive();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_QUADRANT WHERE QUADRANT_NAME = :name", nativeQuery = true)
    Optional<Quadrant> findByName(@Param("name") String name);
}
