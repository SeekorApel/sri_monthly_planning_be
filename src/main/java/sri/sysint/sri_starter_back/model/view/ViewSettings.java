package sri.sysint.sri_starter_back.model.view;

import java.math.BigDecimal;

public class ViewSettings {

    private BigDecimal settingId;
    private String settingKey;
    private String settingValue;
    
    
	public ViewSettings(BigDecimal settingId, String settingKey, String settingValue) {
		super();
		this.settingId = settingId;
		this.settingKey = settingKey;
		this.settingValue = settingValue;
	}
	public BigDecimal getSettingId() {
		return settingId;
	}
	public void setSettingId(BigDecimal settingId) {
		this.settingId = settingId;
	}
	public String getSettingKey() {
		return settingKey;
	}
	public void setSettingKey(String settingKey) {
		this.settingKey = settingKey;
	}
	public String getSettingValue() {
		return settingValue;
	}
	public void setSettingValue(String settingValue) {
		this.settingValue = settingValue;
	}
    
    
    
	
}
