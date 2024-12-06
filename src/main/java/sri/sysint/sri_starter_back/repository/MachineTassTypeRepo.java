package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.model.Plant;


public interface MachineTassTypeRepo extends JpaRepository<MachineTassType, String>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINETASSTYPE WHERE MACHINETASSTYPE_ID = :id", nativeQuery = true)
    Optional<MachineTassType> findById(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINETASSTYPE ORDER BY MACHINETASSTYPE_ID ASC", nativeQuery = true)
    List<MachineTassType> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINETASSTYPE WHERE STATUS = 1", nativeQuery = true)
	List<MachineTassType> findMachineTassTypeActive();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINETASSTYPE WHERE DESCRIPTION = :description", nativeQuery = true)
	Optional<MachineTassType> findByDescription(@Param("description") String description);

}
