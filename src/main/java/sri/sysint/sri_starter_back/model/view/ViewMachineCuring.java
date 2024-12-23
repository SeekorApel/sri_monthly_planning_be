package sri.sysint.sri_starter_back.model.view;

public class ViewMachineCuring {
	
    private String operationShortText;
    
	public ViewMachineCuring(String operationShortText, String wip) {
		super();
		this.operationShortText = operationShortText;
	}
	public ViewMachineCuring() {}
	public String getOperationShortText() {
		return operationShortText;
	}
	public void setOperationShortText(String operationShortText) {
		this.operationShortText = operationShortText;
	}

}
