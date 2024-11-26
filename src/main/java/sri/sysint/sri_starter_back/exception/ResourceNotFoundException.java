package sri.sysint.sri_starter_back.exception;

public class ResourceNotFoundException extends Exception{

	private static final long serialVersionUID = -3684310527743549592L;
	
	public ResourceNotFoundException() {
		super();
	}

	public ResourceNotFoundException(final String message) {
		super(message);
	}

}
