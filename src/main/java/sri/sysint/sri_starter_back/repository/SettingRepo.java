package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Setting;

public interface SettingRepo extends JpaRepository<Setting, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_SETTING WHERE SETTING_ID = :id", nativeQuery = true)
    Optional<Setting> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_SETTING ORDER BY SETTING_ID ASC", nativeQuery = true)
    List<Setting> getDataOrderId();
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_SETTING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_SETTING WHERE STATUS = 1", nativeQuery = true)
	List<Setting> findSettingActive();
	
    @Query(value = "SELECT * FROM SRI_IMPP_M_SETTING WHERE SETTING_VALUE = :settingValue", nativeQuery = true)
    Optional<Setting> findBySettingValue(@Param("settingValue") String settingValue);
    
	@Query(value = "SELECT SETTING_VALUE \r\n"
			+ "FROM SRI_IMPP_M_SETTING \r\n"
			+ "WHERE SETTING_KEY = 'Capacity' AND STATUS = 1", nativeQuery = true)
	String getCapacity();
	
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_SETTING WHERE SETTING_KEY = 'Small Order Limit'", nativeQuery = true)
    Setting findSmallOrderLimit();
}
