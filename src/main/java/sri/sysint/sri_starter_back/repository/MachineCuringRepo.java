package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineTass;
import sri.sysint.sri_starter_back.model.Plant;


public interface MachineCuringRepo extends JpaRepository<MachineCuring, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_CURING WHERE WORK_CENTER_TEXT = :id", nativeQuery = true)
    Optional<MachineCuring> findById(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_CURING ORDER BY WORK_CENTER_TEXT ASC", nativeQuery = true)
    List<MachineCuring> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_MACHINE_CURING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_CURING WHERE STATUS = 1", nativeQuery = true)
	List<MachineCuring> findMachineCuringActive();
}
