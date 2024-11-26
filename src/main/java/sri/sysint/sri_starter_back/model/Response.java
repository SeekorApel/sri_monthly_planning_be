package sri.sysint.sri_starter_back.model;

import java.util.Date;


import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;


/**
 * Response
 */
public class Response {

    private Date timestamp;
    private Integer status;
    private String error;
    private String message;
    private String path;
    private Object data;

    public Response(Date timestamp,Integer status,String error,String message,String path,Object data){
        super();
        this.timestamp = timestamp;
        this.status = status;
        this.error = error;
        this.message = message;
        this.path = path;
        this.data = data;
    }
    
    public Response() {};

    
    public Date getTimestamp() {
		return timestamp;
	}

	public void setTimestamp(Date timestamp) {
		this.timestamp = timestamp;
	}

	public Integer getStatus() {
		return status;
	}

	public void setStatus(Integer status) {
		this.status = status;
	}

	public String getError() {
		return error;
	}

	public void setError(String error) {
		this.error = error;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public Object getData() {
		return data;
	}

	public void setData(Object data) {
		this.data = data;
	}

	public String res(Date timestamp,Integer status,String error,String message,String details,Object data) {
		String value = "";
		ObjectMapper mapper = new ObjectMapper();
		try {
			value = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(new Response(timestamp, status, error, message,details,data));
		} catch (JsonProcessingException e) {
			e.printStackTrace();
		}
		return value;
	}
    
}