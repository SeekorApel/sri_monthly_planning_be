package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineExtruding;
import sri.sysint.sri_starter_back.model.Plant;

public interface MachineExtrudingRepo extends JpaRepository<MachineExtruding, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_EXTRUDING WHERE ID_MACHINE_EXT = :id", nativeQuery = true)
    Optional<MachineExtruding> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_EXTRUDING ORDER BY ID_MACHINE_EXT ASC", nativeQuery = true)
    List<MachineExtruding> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_MACHINE_EXTRUDING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_MACHINE_EXTRUDING WHERE STATUS = 1", nativeQuery = true)
	List<MachineExtruding> findMachineExtrudingActive();
}
