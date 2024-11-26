package sri.sysint.sri_starter_back.security;

import static com.auth0.jwt.algorithms.Algorithm.HMAC512;
import static sri.sysint.sri_starter_back.security.SecurityConstants.EXPIRATION_TIME;
import static sri.sysint.sri_starter_back.security.SecurityConstants.HEADER_STRING;
import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;
import static sri.sysint.sri_starter_back.security.SecurityConstants.TOKEN_PREFIX;
import static sri.sysint.sri_starter_back.security.SecurityConstants.USER_NAME;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

import javax.servlet.FilterChain;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.AuthenticationException;
import org.springframework.security.core.userdetails.User;
import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;

import com.auth0.jwt.JWT;
import com.fasterxml.jackson.databind.ObjectMapper;

import sri.sysint.sri_starter_back.model.Users;

/**
 * JWTAuthFilter
 */
public class JWTAuthenFilter extends UsernamePasswordAuthenticationFilter{
    
    private AuthenticationManager authenticationManager;
    
    public JWTAuthenFilter(AuthenticationManager authenticationManager){
        this.authenticationManager = authenticationManager;
    }

    public Authentication attemptAuthentication(HttpServletRequest req, HttpServletResponse res) throws AuthenticationException{
		try{
            Users creds = new ObjectMapper()
                    .readValue(req.getInputStream(), Users.class);
            
            return authenticationManager.authenticate(
                new UsernamePasswordAuthenticationToken(
                    creds.getUserName(), 
                    creds.getPassword(), 
                    new ArrayList<>())
            );
        }catch(IOException e){
            throw new RuntimeException(e);
        }

    }

    @Override
    protected void successfulAuthentication(HttpServletRequest req, 
                                            HttpServletResponse res,
                                            FilterChain chain,
                                            Authentication auth) throws IOException, ServletException{

        String token = JWT.create()
                .withSubject(((User)auth.getPrincipal()).getUsername())
                .withExpiresAt(new Date(System.currentTimeMillis() + EXPIRATION_TIME))
                .sign(HMAC512(SECRET.getBytes()));
        
        res.addHeader("Access-Control-Expose-Headers", "Authorization");
        res.addHeader("Access-Control-Allow-Headers", "Authorization, X-PINGOTHER, Origin, X-Requested-With, Content-Type, Accept, X-Custom-header");
        res.addHeader(HEADER_STRING, TOKEN_PREFIX + token);
        res.addHeader(USER_NAME, ((User)auth.getPrincipal()).getUsername());
        
        
        
        
        }

    
}