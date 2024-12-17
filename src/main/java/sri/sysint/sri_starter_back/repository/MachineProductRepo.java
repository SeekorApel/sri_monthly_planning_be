package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.transaction.annotation.Transactional;

import sri.sysint.sri_starter_back.model.MachineProduct;

public interface MachineProductRepo extends JpaRepository<MachineProduct, BigDecimal>{
	 	@Modifying
	    @Transactional
	    @Query(value = "INSERT INTO SRI_IMPP_M_MACHINE_PRODUCT (PART_NUMBER, WORK_CENTER_TEXT) " +
	                   "VALUES (:partnum, :wct)", nativeQuery = true)
	    int insertNew(
	        @Param("partnum") BigDecimal partnum,
	        @Param("wct") String wct
	    );
	    
		@Modifying
	    @Transactional 
	    @Query(value = "DELETE FROM SRI_IMPP_M_MACHINE_PRODUCT", nativeQuery = true)
	    void deleteAll();
}
