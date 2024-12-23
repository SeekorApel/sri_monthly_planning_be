package sri.sysint.sri_starter_back.service;

import java.util.ArrayList;
import java.util.List;

import javax.transaction.Transactional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.MachineProduct;
import sri.sysint.sri_starter_back.repository.MachineProductRepo;

@Service
@Transactional
public class MachineProductServiceImpl {
	@Autowired
    private MachineProductRepo machineProductRepo;

	public List<MachineProduct> save(List<MachineProduct> List) {
        List<MachineProduct> saved = new ArrayList<>();
        try {
            for (MachineProduct temp : List) {
                if (temp.getPART_NUMBER() == null) {
                    throw new IllegalArgumentException("Partnumber cannot be null");
                }

                machineProductRepo.insertNew(
                    temp.getPART_NUMBER(),
                    temp.getWORK_CENTER_TEXT()            
                );

                saved.add(temp); 
            }
        } catch (Exception e) {
            System.err.println("Error saving list: " + e.getMessage());
            throw e;
        }
        return saved;
    }
	
    public void deleteAll() {
    	machineProductRepo.deleteAll();
    }
}
