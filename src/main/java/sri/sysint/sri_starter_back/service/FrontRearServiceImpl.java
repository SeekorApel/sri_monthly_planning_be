package sri.sysint.sri_starter_back.service;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.transaction.Transactional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.FrontRear;
import sri.sysint.sri_starter_back.repository.FrontRearRepo;

@Service
@Transactional
public class FrontRearServiceImpl {

    @Autowired
    private FrontRearRepo frontRearRepo;

    public FrontRearServiceImpl(FrontRearRepo frontRearRepo) {
        this.frontRearRepo = frontRearRepo;
    }

    public BigDecimal getNewId() {
        return frontRearRepo.getNextId();
    }

    public List<FrontRear> getAllFrontRear() {
        Iterable<FrontRear> frontRears = frontRearRepo.findAll();
        List<FrontRear> frontRearList = new ArrayList<>();
        for (FrontRear item : frontRears) {
            frontRearList.add(item);
        }
        return frontRearList;
    }

    public List<FrontRear> getFrontRearById(BigDecimal id) {
        return frontRearRepo.findListById(id);
    }

    public List<FrontRear> saveFrontRearList(List<FrontRear> frontRearList) {
        List<FrontRear> savedFrontRears = new ArrayList<>();
        try {
            for (FrontRear frontRear : frontRearList) {
                if (frontRear.getID_FRONT_REAR() == null) {
                    throw new IllegalArgumentException("ID_FRONT_REAR cannot be null");
                }

                frontRearRepo.insertNew(
                    frontRear.getID_FRONT_REAR(),
                    frontRear.getDETAIL_ID_MO(),
                    BigDecimal.valueOf(1), 
                    new Date(),            
                    new Date()             
                );

                savedFrontRears.add(frontRear); 
            }
        } catch (Exception e) {
            System.err.println("Error saving FrontRear list: " + e.getMessage());
            throw e;
        }
        return savedFrontRears;
    }



    public FrontRear updateFrontRear(FrontRear frontRear) {
        try {
            Optional<FrontRear> currentFrontRearOpt = frontRearRepo.findById(frontRear.getID_FRONT_REAR());
            if (currentFrontRearOpt.isPresent()) {
                FrontRear currentFrontRear = currentFrontRearOpt.get();

                currentFrontRear.setDETAIL_ID_MO(frontRear.getDETAIL_ID_MO());
                currentFrontRear.setLAST_UPDATE_DATE(new Date());
                currentFrontRear.setLAST_UPDATED_BY(frontRear.getLAST_UPDATED_BY());

                return frontRearRepo.save(currentFrontRear);
            } else {
                throw new RuntimeException("FrontRear with ID " + frontRear.getID_FRONT_REAR() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating FrontRear: " + e.getMessage());
            throw e;
        }
    }

   
    public List<Map<String, Object>> getAllMarketingOrders() {
        return frontRearRepo.findAllDetailFrMarketingOrders();
    }

    public List<Map<String, Object>> getMarketingOrdersByFrontRearId(BigDecimal idFrontRear) {
        return frontRearRepo.findMarketingOrderByFrontRearId(idFrontRear);
    }

    public void deleteAllFrontRear() {
        frontRearRepo.deleteAllFr();
    }
}
