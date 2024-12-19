package sri.sysint.sri_starter_back.repository;

import java.util.Date;
import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import sri.sysint.sri_starter_back.model.Users;

public interface UserRepo extends JpaRepository<Users, Long> {
	
	@Query(value="SELECT * FROM SRI_APP_USERS WHERE USERNAME = :userName", nativeQuery = true)
	Users findByUserName(String userName);
	
	@Query("SELECT X FROM Users X WHERE X.userName = :userName AND X.resetToken = :resetToken")
	Users findByUserNameAndToken(@Param("userName") String userName, @Param("resetToken") String token);

	@Query(value="SELECT SYSDATE FROM DUAL", nativeQuery = true)
	Date getSysdate();
	
	@Query(value="SELECT ADD_MONTHS(TRUNC(SYSDATE),-14) FROM DUAL", nativeQuery = true)
	Date getSysdateRunPrc();
	
	@Query(value="SELECT TO_CHAR(ADD_MONTHS(SYSDATE,-14),'MON-YYYY') FROM DUAL", nativeQuery = true)
	String getMonthRunPrc();
	
	@Query(value="SELECT TRUNC(SYSDATE) FROM DUAL", nativeQuery = true)
	Date getTruncSysdate();
	
	@Query(value="SELECT * FROM SRI_APP_USERS WHERE USERNAME = :userName AND PASS = :passWord", nativeQuery = true)
	Users findByUsernamePassword(@Param("userName") String userName,@Param("passWord") String passWord);


}
