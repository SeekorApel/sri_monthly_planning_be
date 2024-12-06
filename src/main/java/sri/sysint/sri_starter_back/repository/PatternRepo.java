package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Pattern;
import sri.sysint.sri_starter_back.model.Plant;

public interface PatternRepo extends JpaRepository<Pattern, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_PATTERN WHERE PATTERN_ID = :id", nativeQuery = true)
    Optional<Pattern> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PATTERN ORDER BY PATTERN_ID ASC", nativeQuery = true)
    List<Pattern> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_PATTERN", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PATTERN WHERE STATUS = 1", nativeQuery = true)
	List<Pattern> findPatternActive();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PATTERN WHERE PATTERN_NAME = :name", nativeQuery = true)
	Optional<Pattern> findByName(@Param("name") String name);

}
