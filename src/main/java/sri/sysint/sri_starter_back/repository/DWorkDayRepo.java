package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.DWorkDay;
import sri.sysint.sri_starter_back.model.WorkDay;

public interface DWorkDayRepo extends JpaRepository<DWorkDay, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD WHERE DETAIL_WD_ID = :id", nativeQuery = true)
    Optional<DWorkDay> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD WHERE TO_DATE(DATE_WD, 'DD-MM-YYYY') = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	List<DWorkDay> findByDate(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_WD WHERE DATE_WD = TO_DATE(:id, 'DD-MM-YYYY')", nativeQuery = true)
	Optional<WorkDay> findByDDateWd(@Param("id") String id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD ORDER BY DETAIL_WD_ID ASC", nativeQuery = true)
    List<DWorkDay> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_D_WD", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_D_WD WHERE STATUS = 1", nativeQuery = true)
	List<DWorkDay> findWorkDayActive();
}
