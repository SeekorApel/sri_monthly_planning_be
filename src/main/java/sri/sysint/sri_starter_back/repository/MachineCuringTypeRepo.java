package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.Plant;


public interface MachineCuringTypeRepo extends JpaRepository<MachineCuringType, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINECURINGTYPE WHERE MACHINECURINGTYPE_ID = :id", nativeQuery = true)
    Optional<MachineCuringType> findById(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINECURINGTYPE ORDER BY MACHINECURINGTYPE_ID ASC", nativeQuery = true)
    List<MachineCuringType> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_MACHINECURINGTYPE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINECURINGTYPE WHERE STATUS = 1", nativeQuery = true)
	List<MachineCuringType> findMachineCuringTypeActive();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINECURINGTYPE WHERE DESCRIPTION = :description", nativeQuery = true)
	Optional<MachineCuringType> findByDescription(@Param("description") String description);

}
