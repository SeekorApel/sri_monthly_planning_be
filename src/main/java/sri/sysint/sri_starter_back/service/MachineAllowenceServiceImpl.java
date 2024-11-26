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

import sri.sysint.sri_starter_back.model.MachineAllowence;
import sri.sysint.sri_starter_back.repository.MachineAllowenceRepo;


@Service
@Transactional
public class MachineAllowenceServiceImpl {
	@Autowired
    private MachineAllowenceRepo machineAllowenceRepo;
	
    public MachineAllowenceServiceImpl(MachineAllowenceRepo machineAllowenceRepo){
        this.machineAllowenceRepo = machineAllowenceRepo;
    }
    
    public BigDecimal getNewId() {
    	return machineAllowenceRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<MachineAllowence> getAllMachineAllowence() {
    	Iterable<MachineAllowence> machineAllowences = machineAllowenceRepo.getDataOrderId();
        List<MachineAllowence> machineAllowenceList = new ArrayList<>();
        for (MachineAllowence item : machineAllowences) {
            MachineAllowence machineAllowenceTemp = new MachineAllowence(item);
            machineAllowenceList.add(machineAllowenceTemp);
        }
        
        return machineAllowenceList;
    }
    
    public Optional<MachineAllowence> getMachineAllowenceById(BigDecimal id) {
    	Optional<MachineAllowence> machineAllowence = machineAllowenceRepo.findById(id);
    	return machineAllowence;
    }
    
    public MachineAllowence saveMachineAllowence(MachineAllowence machineAllowence) {
        try {
        	machineAllowence.setMACHINE_ALLOW_ID(getNewId());
            machineAllowence.setSTATUS(BigDecimal.valueOf(1));
            machineAllowence.setCREATION_DATE(new Date());
            machineAllowence.setLAST_UPDATE_DATE(new Date());
            return machineAllowenceRepo.save(machineAllowence);
        } catch (Exception e) {
            System.err.println("Error saving machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineAllowence updateMachineAllowence(MachineAllowence machineAllowence) {
        try {
            Optional<MachineAllowence> currentMachineAllowenceOpt = machineAllowenceRepo.findById(machineAllowence.getMACHINE_ALLOW_ID());
            
            if (currentMachineAllowenceOpt.isPresent()) {
                MachineAllowence currentMachineAllowence = currentMachineAllowenceOpt.get();
                
                currentMachineAllowence.setID_MACHINE(machineAllowence.getID_MACHINE());
                currentMachineAllowence.setPERSON_RESPONSIBLE(machineAllowence.getPERSON_RESPONSIBLE());
                currentMachineAllowence.setSHIFT_1(machineAllowence.getSHIFT_1());
                currentMachineAllowence.setSHIFT_2(machineAllowence.getSHIFT_2());
                currentMachineAllowence.setSHIFT_3(machineAllowence.getSHIFT_3());
                currentMachineAllowence.setSHIFT_1_FRIDAY(machineAllowence.getSHIFT_1_FRIDAY());
                currentMachineAllowence.setTOTAL_SHIFT_123(machineAllowence.getTOTAL_SHIFT_123());
                currentMachineAllowence.setLAST_UPDATE_DATE(new Date());
                currentMachineAllowence.setLAST_UPDATED_BY(machineAllowence.getLAST_UPDATED_BY());
                
                return machineAllowenceRepo.save(currentMachineAllowence);
            } else {
                throw new RuntimeException("MachineAllowence with ID " + machineAllowence.getMACHINE_ALLOW_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineAllowence deleteMachineAllowence(MachineAllowence machineAllowence) {
        try {
            Optional<MachineAllowence> currentMachineAllowenceOpt = machineAllowenceRepo.findById(machineAllowence.getMACHINE_ALLOW_ID());
            
            if (currentMachineAllowenceOpt.isPresent()) {
                MachineAllowence currentMachineAllowence = currentMachineAllowenceOpt.get();
                
                currentMachineAllowence.setSTATUS(BigDecimal.valueOf(0));
                currentMachineAllowence.setLAST_UPDATE_DATE(new Date());
                currentMachineAllowence.setLAST_UPDATED_BY(machineAllowence.getLAST_UPDATED_BY());
                
                return machineAllowenceRepo.save(currentMachineAllowence);
            } else {
                throw new RuntimeException("MachineAllowence with ID " + machineAllowence.getMACHINE_ALLOW_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllMachineAllowence() {
    	machineAllowenceRepo.deleteAll();
    }
    
    public MachineAllowence restoreMachineAllowence(MachineAllowence machineAllowence) {
        try {
            Optional<MachineAllowence> currentMachineAllowenceOpt = machineAllowenceRepo.findById(machineAllowence.getMACHINE_ALLOW_ID());
            
            if (currentMachineAllowenceOpt.isPresent()) {
                MachineAllowence currentMachineAllowence = currentMachineAllowenceOpt.get();
                
                currentMachineAllowence.setSTATUS(BigDecimal.valueOf(1));
                currentMachineAllowence.setLAST_UPDATE_DATE(new Date());
                currentMachineAllowence.setLAST_UPDATED_BY(machineAllowence.getLAST_UPDATED_BY());
                
                return machineAllowenceRepo.save(currentMachineAllowence);
            } else {
                throw new RuntimeException("MachineAllowence with ID " + machineAllowence.getMACHINE_ALLOW_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public ByteArrayInputStream exportMachineAllowenceExcel() throws IOException {
        List<MachineAllowence> machineAllowences = getAllMachineAllowence(); // Mengambil semua data MachineAllowence
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineAllowences); // Konversi data ke Excel
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineAllowence> machineAllowences) throws IOException {
        String[] header = {
            "NOMOR",
            "MACHINE_ALLOW_ID",
            "ID_MACHINE",
            "PERSON_RESPONSIBLE",
            "SHIFT_1",
            "SHIFT_2",
            "SHIFT_3",
            "SHIFT_1_FRIDAY",
            "TOTAL_SHIFT_123"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("MACHINE ALLOWENCE DATA");
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
            for (MachineAllowence m : machineAllowences) {
                Row dataRow = sheet.createRow(rowIndex++);
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(m.getMACHINE_ALLOW_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell idMachineCell = dataRow.createCell(2);
                idMachineCell.setCellValue(m.getID_MACHINE() != null ? m.getID_MACHINE() : null);
                idMachineCell.setCellStyle(borderStyle);

                Cell personResponsibleCell = dataRow.createCell(3);
                personResponsibleCell.setCellValue(m.getPERSON_RESPONSIBLE() != null ? m.getPERSON_RESPONSIBLE() : "");
                personResponsibleCell.setCellStyle(borderStyle);

                Cell shift1Cell = dataRow.createCell(4);
                shift1Cell.setCellValue(m.getSHIFT_1() != null ? m.getSHIFT_1().doubleValue() : null);
                shift1Cell.setCellStyle(borderStyle);

                Cell shift2Cell = dataRow.createCell(5);
                shift2Cell.setCellValue(m.getSHIFT_2() != null ? m.getSHIFT_2().doubleValue() : null);
                shift2Cell.setCellStyle(borderStyle);
                
                Cell shift3Cell = dataRow.createCell(6);
                shift3Cell.setCellValue(m.getSHIFT_3() != null ? m.getSHIFT_3().doubleValue() : null);
                shift3Cell.setCellStyle(borderStyle);

                Cell shift1FridayCell = dataRow.createCell(7);
                shift1FridayCell.setCellValue(m.getSHIFT_1_FRIDAY() != null ? m.getSHIFT_1_FRIDAY().doubleValue() : null);
                shift1FridayCell.setCellStyle(borderStyle);

                Cell totalShiftCell = dataRow.createCell(8);
                totalShiftCell.setCellValue(m.getTOTAL_SHIFT_123() != null ? m.getTOTAL_SHIFT_123().doubleValue() : null);
                totalShiftCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Machine Allowence");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
