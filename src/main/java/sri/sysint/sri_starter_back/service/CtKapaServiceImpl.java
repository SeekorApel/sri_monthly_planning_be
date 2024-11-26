package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.CtKapa;
import sri.sysint.sri_starter_back.repository.CtKapaRepo;


@Service
@Transactional
public class CtKapaServiceImpl {
	@Autowired
    private CtKapaRepo ctKapaRepo;
	
    public CtKapaServiceImpl(CtKapaRepo ctKapaRepo){
        this.ctKapaRepo = ctKapaRepo;
    }
    
    public BigDecimal getNewId() {
    	return ctKapaRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<CtKapa> getAllCtKapa() {
    	Iterable<CtKapa> ctKapas = ctKapaRepo.getDataOrderId();
        List<CtKapa> ctKapaList = new ArrayList<>();
        for (CtKapa item : ctKapas) {
            CtKapa ctKapaTemp = new CtKapa(item);
            ctKapaList.add(ctKapaTemp);
        }
        
        return ctKapaList;
    }
    
    public Optional<CtKapa> getCtKapaById(BigDecimal id) {
    	Optional<CtKapa> ctKapa = ctKapaRepo.findById(id);
    	return ctKapa;
    }
    
    public CtKapa saveCtKapa(CtKapa ctKapa) {
        try {
        	ctKapa.setID_CT_KAPA(getNewId());
            ctKapa.setSTATUS(BigDecimal.valueOf(1));
            ctKapa.setCREATION_DATE(new Date());
            ctKapa.setLAST_UPDATE_DATE(new Date());
            return ctKapaRepo.save(ctKapa);
        } catch (Exception e) {
            System.err.println("Error saving ctKapa: " + e.getMessage());
            throw e;
        }
    }
    
    public CtKapa updateCtKapa(CtKapa ctKapa) {
        try {
            Optional<CtKapa> currentCtKapaOpt = ctKapaRepo.findById(ctKapa.getID_CT_KAPA());
            
            if (currentCtKapaOpt.isPresent()) {
                CtKapa currentCtKapa = currentCtKapaOpt.get();
                
                currentCtKapa.setITEM_CURING(ctKapa.getITEM_CURING());
                currentCtKapa.setTYPE_CURING(ctKapa.getTYPE_CURING());
                currentCtKapa.setDESCRIPTION(ctKapa.getDESCRIPTION());
                currentCtKapa.setCYCLE_TIME(ctKapa.getCYCLE_TIME());
                currentCtKapa.setSHIFT(ctKapa.getSHIFT());
                currentCtKapa.setKAPA_PERSHIFT(ctKapa.getKAPA_PERSHIFT());
                currentCtKapa.setLAST_UPDATE_DATA(ctKapa.getLAST_UPDATE_DATA());
                currentCtKapa.setMACHINE(ctKapa.getMACHINE());
                currentCtKapa.setLAST_UPDATE_DATE(new Date());
                currentCtKapa.setLAST_UPDATED_BY(ctKapa.getLAST_UPDATED_BY());
                
                return ctKapaRepo.save(currentCtKapa);
            } else {
                throw new RuntimeException("CtKapa with ID " + ctKapa.getID_CT_KAPA() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating ctKapa: " + e.getMessage());
            throw e;
        }
    }
    
    public CtKapa deleteCtKapa(CtKapa ctKapa) {
        try {
            Optional<CtKapa> currentCtKapaOpt = ctKapaRepo.findById(ctKapa.getID_CT_KAPA());
            
            if (currentCtKapaOpt.isPresent()) {
                CtKapa currentCtKapa = currentCtKapaOpt.get();
                
                currentCtKapa.setSTATUS(BigDecimal.valueOf(0));
                currentCtKapa.setLAST_UPDATE_DATE(new Date());
                currentCtKapa.setLAST_UPDATED_BY(ctKapa.getLAST_UPDATED_BY());
                
                return ctKapaRepo.save(currentCtKapa);
            } else {
                throw new RuntimeException("CtKapa with ID " + ctKapa.getID_CT_KAPA() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating ctKapa: " + e.getMessage());
            throw e;
        }
    }
    
    public CtKapa restoreCtKapa(CtKapa ctKapa) {
        try {
            Optional<CtKapa> currentCtKapaOpt = ctKapaRepo.findById(ctKapa.getID_CT_KAPA());
            
            if (currentCtKapaOpt.isPresent()) {
                CtKapa currentCtKapa = currentCtKapaOpt.get();
                
                currentCtKapa.setSTATUS(BigDecimal.valueOf(1)); 
                currentCtKapa.setLAST_UPDATE_DATE(new Date());
                currentCtKapa.setLAST_UPDATED_BY(ctKapa.getLAST_UPDATED_BY());
                
                return ctKapaRepo.save(currentCtKapa);
            } else {
                throw new RuntimeException("CtKapa with ID " + ctKapa.getID_CT_KAPA() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring ctKapa: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllCtKapa() {
    	ctKapaRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportCtKapasExcel() throws IOException {
        List<CtKapa> ctKapas = ctKapaRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(ctKapas);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<CtKapa> ctKapas) throws IOException {
        String[] header = {
            "NOMOR",
            "ID_CT_KAPA",
            "ITEM_CURING",
            "TYPE_CURING",
            "DESCRIPTION",
            "CYCLE_TIME",
            "SHIFT",
            "KAPA_PERSHIFT",
            "LAST_UPDATE_DATA",
            "MACHINE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("CT KAPA DATA");
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 1;
            for (CtKapa ctKapa : ctKapas) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCtKapaCell = dataRow.createCell(1);
                idCtKapaCell.setCellValue(ctKapa.getID_CT_KAPA() != null ? ctKapa.getID_CT_KAPA().doubleValue() : null);
                idCtKapaCell.setCellStyle(borderStyle);

                Cell itemCuringCell = dataRow.createCell(2);
                itemCuringCell.setCellValue(ctKapa.getITEM_CURING() != null ? ctKapa.getITEM_CURING() : "");
                itemCuringCell.setCellStyle(borderStyle);

                Cell typeCuringCell = dataRow.createCell(3);
                typeCuringCell.setCellValue(ctKapa.getTYPE_CURING() != null ? ctKapa.getTYPE_CURING() : "");
                typeCuringCell.setCellStyle(borderStyle);

                Cell descriptionCell = dataRow.createCell(4);
                descriptionCell.setCellValue(ctKapa.getDESCRIPTION() != null ? ctKapa.getDESCRIPTION() : "");
                descriptionCell.setCellStyle(borderStyle);

                Cell cycleTimeCell = dataRow.createCell(5);
                cycleTimeCell.setCellValue(ctKapa.getCYCLE_TIME() != null ? ctKapa.getCYCLE_TIME().doubleValue() : null);
                cycleTimeCell.setCellStyle(borderStyle);

                Cell shiftCell = dataRow.createCell(6);
                shiftCell.setCellValue(ctKapa.getSHIFT() != null ? ctKapa.getSHIFT().doubleValue() : null);
                shiftCell.setCellStyle(borderStyle);

                Cell kapaPerShiftCell = dataRow.createCell(7);
                kapaPerShiftCell.setCellValue(ctKapa.getKAPA_PERSHIFT() != null ? ctKapa.getKAPA_PERSHIFT().doubleValue() : null);
                kapaPerShiftCell.setCellStyle(borderStyle);

                Cell lastUpdateDataCell = dataRow.createCell(8);
                lastUpdateDataCell.setCellValue(ctKapa.getLAST_UPDATE_DATA() != null ? ctKapa.getLAST_UPDATE_DATA().doubleValue() : null);
                lastUpdateDataCell.setCellStyle(borderStyle);

                Cell machineCell = dataRow.createCell(9);
                machineCell.setCellValue(ctKapa.getMACHINE() != null ? ctKapa.getMACHINE() : "");
                machineCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export CtKapa data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
