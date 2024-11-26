package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Kapa;

public interface KapaRepo extends JpaRepository<Kapa, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_CT_KAPA WHERE PART_NUMBER = :id", nativeQuery = true)
    Optional<Kapa> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_CT_KAPA", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_CT_KAPA WHERE STATUS = 1", nativeQuery = true)
	List<Kapa> findKapaActive();
}
