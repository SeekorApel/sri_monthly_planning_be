package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.MachineAllowence;
import sri.sysint.sri_starter_back.model.Plant;

public interface MachineAllowenceRepo extends JpaRepository<MachineAllowence, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_ALLOWENCE WHERE MACHINE_ALLOW_ID = :id", nativeQuery = true)
    Optional<MachineAllowence> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_ALLOWENCE ORDER BY MACHINE_ALLOW_ID ASC", nativeQuery = true)
    List<MachineAllowence> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_MACHINE_ALLOWENCE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_ALLOWENCE WHERE STATUS = 1", nativeQuery = true)
	List<MachineAllowence> findMachineAllowenceActive();
}
