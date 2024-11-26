package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.DWorkDayHours;

public interface DWorkDayHoursRepo extends JpaRepository<DWorkDayHours, Date> {
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS WHERE DETAIL_WD_HOURS_ID = :id", nativeQuery = true)
    Optional<DWorkDayHours> findById(@Param("id") BigDecimal id);

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS WHERE TRUNC(DATE_WD) = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	Optional<DWorkDayHours> findDWdHoursByDate(@Param("id") String id);

    @Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS ORDER BY DATE_WD ASC", nativeQuery = true)
    List<DWorkDayHours> getDataOrderByDateDWd();
    
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS WHERE STATUS = 1", nativeQuery = true)
	List<DWorkDayHours> findDWorkDayHoursActive();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_WD_HOURS", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS WHERE DATE_WD = TO_DATE(:date, 'DD-MM-YYYY') AND DESCRIPTION = :description", nativeQuery = true)
	Optional<DWorkDayHours> findDWdHoursByDateAndDescription(@Param("date") String date, @Param("description") String description);

}
