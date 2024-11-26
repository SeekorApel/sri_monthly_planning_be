package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.CTAssy;
import sri.sysint.sri_starter_back.model.CTCuring;

public interface CTCuringRepo extends JpaRepository<CTCuring, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING WHERE CT_CURING_ID = :id", nativeQuery = true)
    Optional<CTCuring> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_CT_CURING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING ORDER BY CT_CURING_ID ASC", nativeQuery = true)
    List<CTCuring> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING WHERE STATUS = 1", nativeQuery = true)
	List<CTCuring> findCtCuringActive();
}
