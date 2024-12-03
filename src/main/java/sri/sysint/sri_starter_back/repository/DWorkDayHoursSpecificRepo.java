package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.DWorkDayHoursSpesific;

public interface DWorkDayHoursSpecificRepo extends JpaRepository<DWorkDayHoursSpesific, Date> {
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE DETAIL_WD_HOURS_SPECIFIC_ID = :id", nativeQuery = true)
    Optional<DWorkDayHoursSpesific> findById(@Param("id") BigDecimal id);

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE TRUNC(DATE_WD) = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	Optional<DWorkDayHoursSpesific> findDWdHoursByDate(@Param("id") String id);

    @Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC ORDER BY DATE_WD ASC", nativeQuery = true)
    List<DWorkDayHoursSpesific> getDataOrderByDateDWd();
    
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE STATUS = 1", nativeQuery = true)
	List<DWorkDayHoursSpesific> findDWorkDayHoursActive();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_WD_HOURS_SPECIFIC", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE TO_DATE(DATE_WD, 'DD-MM-YYYY') = TO_DATE(:date, 'DD-MM-YYYY') AND DESCRIPTION = :description", nativeQuery = true)
	Optional<DWorkDayHoursSpesific> findDWdHoursByDateAndDescription(@Param("date") String date, @Param("description") String description);

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC", nativeQuery = true)
	List<DWorkDayHoursSpesific> fetchAllManual();

	@Query(value = "SELECT * FROM SRI_IMPP_D_WD_HOURS_SPECIFIC WHERE EXTRACT(MONTH FROM DATE_WD) = :month AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
	List<DWorkDayHoursSpesific> findDWdHoursByMonthAndYear(@Param("month") int month, @Param("year") int year);

}
