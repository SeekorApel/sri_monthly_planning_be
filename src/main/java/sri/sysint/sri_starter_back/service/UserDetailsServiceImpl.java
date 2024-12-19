package sri.sysint.sri_starter_back.service;

import static java.util.Collections.emptyList;

import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.userdetails.User;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Roles;
import sri.sysint.sri_starter_back.model.Users;
import sri.sysint.sri_starter_back.repository.RoleRepo;
import sri.sysint.sri_starter_back.repository.UserRepo;


/**
 * UserDetailsServiceImpl
 */
@Service
public class UserDetailsServiceImpl implements UserDetailsService {
	
	@Autowired
    private UserRepo userRepo;
	
	@Autowired
	private RoleRepo roleRepo;
	
    public UserDetailsServiceImpl(UserRepo userRepo){
        this.userRepo = userRepo;
    }
    
    public Users getUserByUsername(String userName) {
        return userRepo.findByUserName(userName);
    }

    @Override
    public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException{
        Users user = userRepo.findByUserName(username);
        if(user == null){
            throw new UsernameNotFoundException(username);
        }
        return new User(user.getUserName(), user.getPassword(), emptyList());
    }
    
    public List<Users> getAllUser() {
        List<Users> users = userRepo.findAll();
        return users.isEmpty() ? null : users;
    }

    public Users getUserByUserName(String userName) {
    	Users user = userRepo.findByUserName(userName);
    	return user;
    }
        
    public Roles getRoleByUserId(Long userId) {
        Roles role = roleRepo.findById(userId).orElse(null);
        System.out.println(role);
        return (role != null) ? role : null;
    }
    
    public Users getUser(String username, String password) {
    	Users user = userRepo.findByUsernamePassword(username, password);
    	if(user != null) {
    		return user;
    	}else {
    		return null;
    	}
    }

       
}