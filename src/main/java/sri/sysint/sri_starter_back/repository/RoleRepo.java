package sri.sysint.sri_starter_back.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import sri.sysint.sri_starter_back.model.Roles;

public interface RoleRepo extends JpaRepository<Roles, Long> {

}
