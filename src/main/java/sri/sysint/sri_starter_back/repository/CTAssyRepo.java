package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.CTAssy;
import sri.sysint.sri_starter_back.model.CuringSize;

public interface CTAssyRepo extends JpaRepository<CTAssy, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_ASSY WHERE CT_ASSY_ID = :id", nativeQuery = true)
    Optional<CTAssy> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_CT_ASSY", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_ASSY ORDER BY CT_ASSY_ID ASC", nativeQuery = true)
    List<CTAssy> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_ASSY WHERE STATUS = 1", nativeQuery = true)
	List<CTAssy> findCtAssyActive();
}
