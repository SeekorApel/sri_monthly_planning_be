package sri.sysint.sri_starter_back.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import sri.sysint.sri_starter_back.model.Dummy;

public interface DummyRepo extends JpaRepository<Dummy, Integer>{

	Dummy findByCode(String code);
	
}
