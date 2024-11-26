package sri.sysint.sri_starter_back.security;

public class SecurityConstants {
    public static final String SECRET = "AERONOORD_SECRET_KEY";
    public static final long EXPIRATION_TIME = 86400000; // 1 days
    public static final String TOKEN_PREFIX = "Bearer ";
    public static final String HEADER_STRING = "Authorization";
    public static final String USER_NAME = "username";
    public static final String SIGN_UP_URL = "/sign-up";
    public static final String GET_USERNAME_URL = "/username/**";
    public static final String HOME = "/home";
    
    public static final String URI_VERSION = "/version/**";
    public static final String URI_SENDTOKEN = "/user/sendtoken/**";
    public static final String URI_USER_RESET = "/user/reset";
    public static final String URI_UPLOAD ="/upload";
    public static final String URI_UNLOCK ="/**";
    
}