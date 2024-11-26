package sri.sysint.sri_starter_back;

import javax.annotation.Resource;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;

import sri.sysint.sri_starter_back.SriStarterBackApplication;
import sri.sysint.sri_starter_back.service.StorageService;

@SpringBootApplication
public class SriStarterBackApplication {

	@Bean
	public BCryptPasswordEncoder bCryptPasswordEncoder(){
		return new BCryptPasswordEncoder();
	}
	
	@Resource
	StorageService storageService;
	public static void main(String[] args) {
		SpringApplication.run(SriStarterBackApplication.class, args);
	}
	
	public void run(String... arg) throws Exception {
		storageService.deleteAll();
		storageService.init();
	}
}
