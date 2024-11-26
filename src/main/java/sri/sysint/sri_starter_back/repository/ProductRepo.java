package sri.sysint.sri_starter_back.repository;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.ProductType;

public interface ProductRepo extends JpaRepository<Product, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCT WHERE PART_NUMBER = :id", nativeQuery = true)
    Optional<Product> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_PRODUCT", nativeQuery = true)
	BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCT ORDER BY PART_NUMBER ASC", nativeQuery = true)
    List<Product> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_PRODUCT WHERE STATUS = 1", nativeQuery = true)
	List<Product> findProductActive();
}
