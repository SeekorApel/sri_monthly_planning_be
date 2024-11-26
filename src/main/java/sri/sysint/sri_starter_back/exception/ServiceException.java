package sri.sysint.sri_starter_back.exception;

public class ServiceException extends Exception{

	private static final long serialVersionUID = -8641348662433235230L;
	
	public ServiceException() {
		super();
	}
	
	public ServiceException(final String message) {
		super(message);
	}

}
