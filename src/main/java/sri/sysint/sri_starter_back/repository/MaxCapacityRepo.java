package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MaxCapacity;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Users;

public interface MaxCapacityRepo extends JpaRepository<MaxCapacity, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_D_MAXCAPACITY WHERE MAX_CAP_ID = :id", nativeQuery = true)
    Optional<MaxCapacity> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_MAXCAPACITY ORDER BY MAX_CAP_ID ASC", nativeQuery = true)
    List<MaxCapacity> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_MAXCAPACITY", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_MAXCAPACITY WHERE STATUS = 1", nativeQuery = true)
	List<MaxCapacity> findMaxCapacityActive();
}
