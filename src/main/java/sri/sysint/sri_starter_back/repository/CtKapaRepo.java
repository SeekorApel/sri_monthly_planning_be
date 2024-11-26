package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.CtKapa;
import sri.sysint.sri_starter_back.model.Plant;

public interface CtKapaRepo extends JpaRepository<CtKapa, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_KAPA WHERE ID_CT_KAPA = :id", nativeQuery = true)
    Optional<CtKapa> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_KAPA ORDER BY ID_CT_KAPA ASC", nativeQuery = true)
    List<CtKapa> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_CT_KAPA", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_KAPA WHERE STATUS = 1", nativeQuery = true)
	List<CtKapa> findCtKapaActive();
}
