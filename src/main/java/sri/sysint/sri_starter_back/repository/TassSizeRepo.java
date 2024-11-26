package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.TassSize;

public interface TassSizeRepo extends JpaRepository<TassSize, BigDecimal>{
    @Query(value = "SELECT * FROM SRI_IMPP_D_TASSSIZE WHERE TASSIZE_ID = :id", nativeQuery = true)
    Optional<TassSize> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_TASSSIZE", nativeQuery = true)
	BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_TASSSIZE ORDER BY TASSIZE_ID ASC", nativeQuery = true)
    List<TassSize> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_TASSSIZE WHERE STATUS = 1", nativeQuery = true)
	List<TassSize> findTassSizeActive();
}
