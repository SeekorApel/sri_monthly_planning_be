package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.StopMachine;

public interface StopMachineRepo extends JpaRepository<StopMachine, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_STOP_MACHINE WHERE STOP_MACHINE_ID = :id", nativeQuery = true)
    Optional<StopMachine> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_STOP_MACHINE ORDER BY STOP_MACHINE_ID ASC", nativeQuery = true)
    List<StopMachine> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_STOP_MACHINE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_STOP_MACHINE WHERE STATUS = 1", nativeQuery = true)
	List<StopMachine> findStopMachineActive();
}
