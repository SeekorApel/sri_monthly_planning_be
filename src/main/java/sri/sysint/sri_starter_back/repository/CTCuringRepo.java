package sri.sysint.sri_starter_back.repository;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.CTAssy;
import sri.sysint.sri_starter_back.model.CTCuring;

public interface CTCuringRepo extends JpaRepository<CTCuring, BigDecimal>{
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING WHERE CT_CURING_ID = :id", nativeQuery = true)
    Optional<CTCuring> findById(@Param("id") BigDecimal id);
	
	@Query(value = "SELECT COUNT(*) FROM SRI_IMPP_M_CT_CURING", nativeQuery = true)
    BigDecimal getNewId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING ORDER BY CT_CURING_ID ASC", nativeQuery = true)
    List<CTCuring> getDataOrderId();
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING WHERE STATUS = 1", nativeQuery = true)
	List<CTCuring> findCtCuringActive();
	
	@Query(value = "SELECT OPERATION_SHORT_TEXT FROM SRI_IMPP_M_CT_CURING WHERE WIP = :itemCuring", nativeQuery = true)
	List<Map<String, Object>> getMachineByItemCuring(@Param("itemCuring") String itemCuring);

	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
			+ "JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "WHERE m.BUILDING_ID = 4 AND WIP = :itemCuring", nativeQuery = true)
    List<CTCuring> findCBuildingMachineByItemCuring(@Param("itemCuring") String itemCuring);

	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
				+ "JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
				+ "WHERE m.BUILDING_ID = 6 AND WIP = :itemCuring", nativeQuery = true)
	    List<CTCuring> findGBuildingMachineByItemCuring(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
				+ "JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
				+ "WHERE m.BUILDING_ID = 6 AND WIP = :itemCuring", nativeQuery = true)
	    List<CTCuring> findDBuildingMachineByItemCuring(@Param("itemCuring") String itemCuring);

	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
			+ "JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "WHERE m.MACHINE_TYPE = 'AB' AND WIP = :itemCuring "
    + "ORDER BY m.CAVITY DESC", nativeQuery = true)
    List<CTCuring> findAirbagMachineByItemCuring(@Param("itemCuring") String itemCuring);

	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
			+ "JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "WHERE m.CAVITY = 2 AND WIP = :itemCuring", nativeQuery = true)
	List<CTCuring> find2CavMachineByItemCuring(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND (c.OPERATION_SHORT_TEXT LIKE '%GD G%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD C%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD D%') "
			+ "			ORDER BY CASE "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD G%' THEN 1 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD C%' THEN 2 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD D%' THEN 3 "
			+ "			ELSE 4 END, "
	  + "     m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineByGCDBuilding(@Param("itemCuring") String partNumber);
	
	@Query(value = "SELECT ct.* FROM SRI_IMPP_M_CT_CURING ct "
			+ "      JOIN SRI_IMPP_M_MACHINE_CURING m ON cT.OPERATION_SHORT_TEXT = m.WORK_CENTER_TEXT "
			+ "			WHERE WIP = itemCuring "
			+ "			AND (OPERATION_SHORT_TEXT LIKE '%GD G%' "
			+ "			OR OPERATION_SHORT_TEXT LIKE '%GD C%' "
			+ "			OR OPERATION_SHORT_TEXT LIKE '%GD D%') "
			+ "      AND m.CAVITY = 2 "
			+ "			ORDER BY CASE "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD G%' THEN 1 "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD C%' THEN 2 "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD D%' THEN 3 "
			+ "			ELSE 4 END", nativeQuery = true)
	List<CTCuring> findMachine2CavByGCDBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c"
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "     AND c.OPERATION_SHORT_TEXT LIKE '%BOM%' "
			+ "			AND (c.OPERATION_SHORT_TEXT LIKE '%GD G%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD C%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD D%') "
			+ "			ORDER BY CASE "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD G%' THEN 1 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD C%' THEN 2 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD D%' THEN 3 "
			+ "			ELSE 4 END, "
	  + "     m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineBOMByGCDBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c"
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "			WHERE c.WIP = :itemCuring "
			+ "     AND (c.OPERATION_SHORT_TEXT LIKE '%AB%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD G%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD C%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD D%') "
			+ "			ORDER BY CASE "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD G%' THEN 1 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD C%' THEN 2 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD D%' THEN 3 "
			+ "			ELSE 4 END, "
	  + "     m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineABByGCDBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c"
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "			WHERE c.WIP = :itemCuring "
			+ "     AND c.OPERATION_SHORT_TEXT LIKE '%GD A%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD B%' "
			+ "			ORDER BY CASE  "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD A%' THEN 1 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD B%' THEN 2 "
			+ "			ELSE 3 END, "
	  + "     m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineByABBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT ct.* FROM SRI_IMPP_M_CT_CURING ct "
			+ "     JOIN SRI_IMPP_M_MACHINE_CURING m ON cT.OPERATION_SHORT_TEXT = m.WORK_CENTER_TEXT "
			+ "			WHERE WIP = itemCuring "
			+ "			AND (OPERATION_SHORT_TEXT LIKE '%GD A%' "
			+ "			OR OPERATION_SHORT_TEXT LIKE '%GD B%') "
			+ "     AND m.CAVITY = 2 "
			+ "			ORDER BY CASE "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD A%' THEN 1 "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD B%' THEN 2 "
			+ "			ELSE 4 END", nativeQuery = true)
	List<CTCuring> findMachine2CavByABBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING  c"
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%BOM%' "
			+ "			AND  ( c.OPERATION_SHORT_TEXT LIKE '%GD A%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD B%') "
			+ "			ORDER BY CASE  "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD A%' THEN 1 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD B%' THEN 2 "
			+ "			ELSE 3 END, "
	  + "     m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineBOMByABBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c"
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%AB%' "
			+ "			AND  ( c.OPERATION_SHORT_TEXT LIKE '%GD A%' "
			+ "			OR c.OPERATION_SHORT_TEXT LIKE '%GD B%') "
			+ "			ORDER BY CASE "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD A%' THEN 1 "
			+ "			WHEN c.OPERATION_SHORT_TEXT LIKE '%GD B%' THEN 2 "
			+ "			ELSE 3 END, "
	  + "     m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineABByABBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%GD H%' "
	  + "     ORDER BY m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineByHBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%BOM%' "
			+ "			AND  ( c.OPERATION_SHORT_TEXT LIKE '%GD H%') "
	  + "     ORDER BY m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineBOMByHBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%AB%' "
			+ "			AND  ( c.OPERATION_SHORT_TEXT LIKE '%GD H%') "
	  + "     ORDER BY m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineABByHBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%AB%' "
	  + "     ORDER BY m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineABAllBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
			+ "			AND c.OPERATION_SHORT_TEXT LIKE '%BOM%' "
	  + "     ORDER BY m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineBOMAllBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT ct.* FROM SRI_IMPP_M_CT_CURING ct "
			+ "      JOIN SRI_IMPP_M_MACHINE_CURING m ON cT.OPERATION_SHORT_TEXT = m.WORK_CENTER_TEXT "
			+ "			WHERE WIP = itemCuring "
			+ "			AND (OPERATION_SHORT_TEXT LIKE '%GD H%') "
			+ "      AND m.CAVITY = 2 "
			+ "			ORDER BY CASE "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD H%' THEN 1 "
			+ "			ELSE 4 END", nativeQuery = true)
	List<CTCuring> findMachine2CavByHBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT ct.* FROM SRI_IMPP_M_CT_CURING ct "
			+ "			JOIN SRI_IMPP_M_MACHINE_CURING m ON cT.OPERATION_SHORT_TEXT = m.WORK_CENTER_TEXT "
			+ "			WHERE WIP = :itemCuring"
			+ "			AND (OPERATION_SHORT_TEXT LIKE '%GD G%' "
			+ "			OR OPERATION_SHORT_TEXT LIKE '%GD C%' "
			+ "			OR OPERATION_SHORT_TEXT LIKE '%GD D%' "
			+ "      OR OPERATION_SHORT_TEXT LIKE '%GD A%' "
			+ "      OR OPERATION_SHORT_TEXT LIKE '%GD B%' "
			+ "      OR OPERATION_SHORT_TEXT LIKE '%GD H%')  "
			+ "			AND m.CAVITY = 2  "
			+ "			ORDER BY CASE  "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD G%' THEN 1  "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD C%' THEN 2  "
			+ "			WHEN OPERATION_SHORT_TEXT LIKE '%GD D%' THEN 3  "
			+ "      WHEN OPERATION_SHORT_TEXT LIKE '%GD A%' THEN 4  "
			+ "      WHEN OPERATION_SHORT_TEXT LIKE '%GD B%' THEN 5  "
			+ "      WHEN OPERATION_SHORT_TEXT LIKE '%GD H%' THEN 6  "
			+ "			ELSE 4 END", nativeQuery = true)
	List<CTCuring> findMachine2CavAllBuilding(@Param("itemCuring") String itemCuring);
	
	@Query(value = "SELECT * FROM SRI_IMPP_M_CT_CURING c "
	  + "     JOIN SRI_IMPP_M_MACHINE_CURING m ON m.WORK_CENTER_TEXT = c.OPERATION_SHORT_TEXT "
			+ "     WHERE c.WIP = :itemCuring "
	  + "     ORDER BY m.CAVITY DESC", nativeQuery = true)
	List<CTCuring> findMachineByItemCuring(@Param("itemCuring") String itemCuring);
	
}
