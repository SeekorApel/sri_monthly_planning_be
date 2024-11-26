package sri.sysint.sri_starter_back.service;

import java.math.BigDecimal;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javax.transaction.Transactional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Kapa;
import sri.sysint.sri_starter_back.repository.KapaRepo;

@Service
@Transactional
public class KapaServiceImpl {
    @Autowired
    private KapaRepo kapaRepo;

    public KapaServiceImpl(KapaRepo kapaRepo) {
        this.kapaRepo = kapaRepo;
    }
    
    public BigDecimal getNewId() {
        return kapaRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<Kapa> getAllKapa() {
        Iterable<Kapa> kapas = kapaRepo.findAll();
        List<Kapa> kapaList = new ArrayList<>();
        for (Kapa item : kapas) {
            Kapa kapaTemp = new Kapa(item);
            kapaList.add(kapaTemp);
        }
        return kapaList;
    }
    
    public Optional<Kapa> getKapaById(BigDecimal id) {
        return kapaRepo.findById(id);
    }
    
    public Kapa saveKapa(Kapa kapa) {
        try {
            kapa.setLAST_UPDATE_DATA(BigDecimal.valueOf(1));
            kapa.setCREATION_DATE(new Date());
            kapa.setLAST_UPDATE_DATE(new Date());
            return kapaRepo.save(kapa);
        } catch (Exception e) {
            System.err.println("Error saving kapa: " + e.getMessage());
            throw e;
        }
    }
    
    public Kapa updateKapa(Kapa kapa) {
        try {
            Optional<Kapa> currentKapaOpt = kapaRepo.findById(kapa.getPART_NUMBER());
            if (currentKapaOpt.isPresent()) {
                Kapa currentKapa = currentKapaOpt.get();
                
                currentKapa.setITEM_CURING(kapa.getITEM_CURING());
                currentKapa.setTYPE_CURING(currentKapa.getITEM_CURING());
                currentKapa.setDESCRIPTION(currentKapa.getDESCRIPTION());
                currentKapa.setCYCLE_TIME(currentKapa.getCYCLE_TIME());
                currentKapa.setSHIFT(currentKapa.getSHIFT());
                currentKapa.setKAPA_PERSHIFT(currentKapa.getKAPA_PERSHIFT());
                currentKapa.setLAST_UPDATE_DATA(currentKapa.getLAST_UPDATE_DATA());
                currentKapa.setMACHINE(currentKapa.getMACHINE());
                
                currentKapa.setLAST_UPDATE_DATE(new Date());
                currentKapa.setLAST_UPDATED_BY(kapa.getLAST_UPDATED_BY());
                return kapaRepo.save(currentKapa);
            } else {
                throw new RuntimeException("Kapa with ID " + kapa.getPART_NUMBER() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating kapa: " + e.getMessage());
            throw e;
        }
    }
    
    public Kapa deleteKapa(Kapa kapa) {
        try {
            Optional<Kapa> currentKapaOpt = kapaRepo.findById(kapa.getPART_NUMBER());
            if (currentKapaOpt.isPresent()) {
                Kapa currentKapa = currentKapaOpt.get();
                currentKapa.setSTATUS(BigDecimal.valueOf(0));
                currentKapa.setLAST_UPDATE_DATE(new Date());
                currentKapa.setLAST_UPDATED_BY(kapa.getLAST_UPDATED_BY());
                return kapaRepo.save(currentKapa);
            } else {
                throw new RuntimeException("Kapa with ID " + kapa.getPART_NUMBER() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating kapa: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllKapa() {
        kapaRepo.deleteAll();
    }
}
