package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Size;

public interface SizeRepo extends JpaRepository<Size, String>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_SIZE WHERE SIZE_ID = :id", nativeQuery = true)
    Optional<Size> findById(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_SIZE ORDER BY SIZE_ID ASC", nativeQuery = true)
    List<Size> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_SIZE WHERE STATUS = 1", nativeQuery = true)
	List<Size> findSizeActive();
}
