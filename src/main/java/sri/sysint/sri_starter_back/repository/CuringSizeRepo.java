package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.CuringSize;
import sri.sysint.sri_starter_back.model.TassSize;

public interface CuringSizeRepo extends JpaRepository<CuringSize, BigDecimal>{
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_CURINGSIZE WHERE CURINGSIZE_ID = :id", nativeQuery = true)
    Optional<CuringSize> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_CURINGSIZE", nativeQuery = true)
	BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_CURINGSIZE ORDER BY CURINGSIZE_ID ASC", nativeQuery = true)
    List<CuringSize> getDataOrderId();

	@Query(value = "SELECT * FROM SRI_IMPP_D_CURINGSIZE WHERE STATUS = 1", nativeQuery = true)
	List<CuringSize> findCuringSizeActive();
}
