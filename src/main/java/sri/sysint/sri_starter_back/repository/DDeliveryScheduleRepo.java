package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.DDeliverySchedule;

public interface DDeliveryScheduleRepo extends JpaRepository<DDeliverySchedule, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_D_DELIVERYSCHEDULE WHERE DETAIL_DS_ID = :id", nativeQuery = true)
    Optional<DDeliverySchedule> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_DELIVERYSCHEDULE ORDER BY DETAIL_DS_ID ASC", nativeQuery = true)
    List<DDeliverySchedule> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_DELIVERYSCHEDULE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_DELIVERYSCHEDULE WHERE STATUS = 1", nativeQuery = true)
	List<DDeliverySchedule> findDDeliveryScheduleActive();
}
