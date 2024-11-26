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
import sri.sysint.sri_starter_back.model.QuadrantDistance;

public interface QuadrantDistanceRepo extends JpaRepository<QuadrantDistance, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_D_QUADRANTDISTANCE WHERE ID_Q_DISTANCE = :id", nativeQuery = true)
    Optional<QuadrantDistance> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_QUADRANTDISTANCE ORDER BY ID_Q_DISTANCE ASC", nativeQuery = true)
    List<QuadrantDistance> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_QUADRANTDISTANCE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_QUADRANTDISTANCE WHERE STATUS = 1", nativeQuery = true)
	List<QuadrantDistance> findQuadrantDistanceActive();
}
