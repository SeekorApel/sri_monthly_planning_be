package sri.sysint.sri_starter_back.repository;

import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.DWorkDayHours;
import sri.sysint.sri_starter_back.model.WorkDay;

public interface WorkDayRepo extends JpaRepository<WorkDay, Date> {
    
	@Query(value = "SELECT * FROM SRI_IMPP_M_WD WHERE TRUNC(DATE_WD) = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	Optional<WorkDay> findByDDateWd(@Param("id") String id);

    @Query(value = "SELECT * FROM SRI_IMPP_M_WD ORDER BY DATE_WD ASC", nativeQuery = true)
    List<WorkDay> getDataOrderByDateWd();
    
	@Query(value = "SELECT * FROM SRI_IMPP_M_WD WHERE STATUS = 1", nativeQuery = true)
	List<WorkDay> findWorkDayActive();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_WD WHERE DATE_WD BETWEEN TO_DATE(:startDate, 'DD-MM-YYYY') AND TO_DATE(:endDate, 'DD-MM-YYYY') ORDER BY DATE_WD ASC", nativeQuery = true)
	List<WorkDay> findAllByDateRange(@Param("startDate") String startDate, @Param("endDate") String endDate);
	
    @Query(value = "SELECT * " +
	        "FROM SRI_IMPP_M_WD " +
	        "WHERE EXTRACT(MONTH FROM DATE_WD) = :month " +
	        "AND EXTRACT(YEAR FROM DATE_WD) = :year", nativeQuery = true)
	List<WorkDay> findByMonthYear(@Param("month") int month,@Param("year") int year);

}
