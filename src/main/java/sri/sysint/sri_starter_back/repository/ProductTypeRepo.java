package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;

import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.ProductType;

public interface ProductTypeRepo extends JpaRepository<ProductType, BigDecimal>{
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCTTYPE WHERE PRODUCT_TYPE_ID = :id", nativeQuery = true)
    Optional<ProductType> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_PRODUCTTYPE", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCTTYPE ORDER BY PRODUCT_TYPE_ID ASC", nativeQuery = true)
    List<ProductType> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCTTYPE WHERE STATUS = 1", nativeQuery = true)
	List<ProductType> findProductTypeActive();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCTTYPE WHERE CATEGORY = :category", nativeQuery = true)
	Optional<ProductType> findByCategory(@Param("category") String category);

}
