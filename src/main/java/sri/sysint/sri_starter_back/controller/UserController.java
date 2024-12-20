package sri.sysint.sri_starter_back.controller;
import java.util.Date;
import java.util.List;
import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.servlet.http.HttpServletRequest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.Roles;
import sri.sysint.sri_starter_back.model.Users;
import sri.sysint.sri_starter_back.model.view.ViewUsers;
import sri.sysint.sri_starter_back.security.SecurityConstants;
import sri.sysint.sri_starter_back.service.UserDetailsServiceImpl;
import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;
import static sri.sysint.sri_starter_back.security.SecurityConstants.HEADER_STRING;
import static sri.sysint.sri_starter_back.security.SecurityConstants.TOKEN_PREFIX;

@CrossOrigin(maxAge = 3600)
@RestController
public class UserController {
		
	private Response response;
	
	@Autowired
	private UserDetailsServiceImpl userDetailsServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
	private boolean validateToken(HttpServletRequest req) throws ResourceNotFoundException {		
	    String header = req.getHeader(HEADER_STRING);
	    if (header == null || !header.startsWith(TOKEN_PREFIX)) {
	    	throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace(TOKEN_PREFIX, "");
	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	                .build()
	                .verify(token)
	                .getSubject();

	        if (user == null) {
	            throw new ResourceNotFoundException("User not found");
	        }
	        
	    } catch (Exception e) {
	        return false;
	    }
	    
	    return true;
	}

	
	//START - GET MAPPING
	@GetMapping("/getAllUser")
	public Response getAllUser(final HttpServletRequest req) throws ResourceNotFoundException {
	    boolean isValidToken = validateToken(req);
	    
	    if (isValidToken) {
	        try {
	            List<Users> users = userDetailsServiceImpl.getAllUser();
	            response = new Response(new Date(), HttpStatus.OK.value(), null, HttpStatus.OK.getReasonPhrase(), req.getRequestURI(), users);
	        } catch (Exception e) {
	            response = new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, e.getMessage(), req.getRequestURI(), null);
	        }
	    } else {
	        response = new Response(new Date(), HttpStatus.UNAUTHORIZED.value(), null, "Invalid or missing token", req.getRequestURI(), null);
	    }

	    return response;
	}

	
	@GetMapping("/getUsername/{userName}")
	public Response getUserByUsername(@PathVariable String userName, HttpServletRequest req) throws ResourceNotFoundException {
		
	    try {
	    	Users user = userDetailsServiceImpl.getUserByUsername(userName);
	    	if(user != null) {
	    		Roles role = userDetailsServiceImpl.getRoleByUserId(user.getId());
	    		if(role != null) {
	    			Roles dataRole = new Roles(role);
	    			ViewUsers dataUser = new ViewUsers(user, dataRole);
	    			response = new Response(new Date(), HttpStatus.OK.value(), null, HttpStatus.OK.getReasonPhrase(), req.getRequestURI(), dataUser);
	    		}
	    	}else {
	    		response = new Response(new Date(), HttpStatus.NOT_FOUND.value(), null, HttpStatus.NOT_FOUND.getReasonPhrase(), req.getRequestURI(), null);
	    	}
	    }catch(Exception e) {
	    	response = new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, e.getMessage(), req.getRequestURI(), null);
	    }

	    return response;
	}
	
	//END - GET MAPPING
	
	//START - POST MAPPING
	@PostMapping("/loginUser")
	public Response signin(final HttpServletRequest req, @RequestBody Users user) throws ResourceNotFoundException {
	    Users existingUser = userDetailsServiceImpl.getUser(user.getUserName(), user.getPassword());

	    if (existingUser != null) {
	        String token = JWT.create()
	                .withSubject(existingUser.getUserName())
	                .withExpiresAt(new Date(System.currentTimeMillis() + SecurityConstants.EXPIRATION_TIME))
	                .sign(Algorithm.HMAC512(SecurityConstants.SECRET.getBytes()));

	        // Include the token in the response
	        response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                token // Pass the generated token in the response
	        );
	    } else {
	        response = new Response(
	                new Date(),
	                HttpStatus.NOT_FOUND.value(),
	                null,
	                "User not found",
	                req.getRequestURI(),
	                null
	        );
	    }

	    return response;
	}
	
	
	//END - POST MAPPING
//START - PUT MAPPING
//END - PUT MAPPING
//START - DELETE MAPPING
//END - DELETE MAPPING
//START - PROCEDURE
//END - PROCEDURE
}
