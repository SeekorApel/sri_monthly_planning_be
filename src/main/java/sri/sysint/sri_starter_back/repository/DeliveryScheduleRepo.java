package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.DeliverySchedule;
import sri.sysint.sri_starter_back.model.Plant;

public interface DeliveryScheduleRepo extends JpaRepository<DeliverySchedule, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_DELIVERYSCHEDULE WHERE DS_ID = :id", nativeQuery = true)
    Optional<DeliverySchedule> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_DELIVERYSCHEDULE ORDER BY DS_ID ASC", nativeQuery = true)
    List<DeliverySchedule> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_DELIVERYSCHEDULE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_DELIVERYSCHEDULE WHERE STATUS = 1", nativeQuery = true)
	List<DeliverySchedule> findDeliveryScheduleActive();
}
