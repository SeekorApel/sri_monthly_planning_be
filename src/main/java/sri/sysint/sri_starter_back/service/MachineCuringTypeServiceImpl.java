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

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;


@Service
@Transactional
public class MachineCuringTypeServiceImpl {
	@Autowired
    private MachineCuringTypeRepo machineCuringTypeRepo;

    public MachineCuringTypeServiceImpl(MachineCuringTypeRepo machineCuringTypeRepo) {
        this.machineCuringTypeRepo = machineCuringTypeRepo;
    }

    public List<MachineCuringType> getAllMachineCuringType() {
        Iterable<MachineCuringType> machineCuringTypes = machineCuringTypeRepo.getDataOrderId();
        List<MachineCuringType> machineCuringTypeList = new ArrayList<>();
        for (MachineCuringType item : machineCuringTypes) {
            MachineCuringType machineCuringTypeTemp = new MachineCuringType(item);
            machineCuringTypeList.add(machineCuringTypeTemp);
        }
        return machineCuringTypeList;
    }

    public Optional<MachineCuringType> getMachineCuringTypeById(String id) {
        Optional<MachineCuringType> machineCuringType = machineCuringTypeRepo.findById(id);
        return machineCuringType;
    }

    public MachineCuringType saveMachineCuringType(MachineCuringType machineCuringType) {
        try {
            machineCuringType.setMACHINECURINGTYPE_ID(machineCuringType.getMACHINECURINGTYPE_ID());
            machineCuringType.setSTATUS(BigDecimal.valueOf(1));
            machineCuringType.setCREATION_DATE(new Date());
            machineCuringType.setLAST_UPDATE_DATE(new Date());
            return machineCuringTypeRepo.save(machineCuringType);
        } catch (Exception e) {
            System.err.println("Error saving machineCuringType: " + e.getMessage());
            throw e;
        }
    }

    public MachineCuringType updateMachineCuringType(MachineCuringType machineCuringType) {
        try {
            Optional<MachineCuringType> currentMachineCuringTypeOpt = machineCuringTypeRepo.findById(machineCuringType.getMACHINECURINGTYPE_ID());
            if (currentMachineCuringTypeOpt.isPresent()) {
                MachineCuringType currentMachineCuringType = currentMachineCuringTypeOpt.get();
                currentMachineCuringType.setSETTING_ID(machineCuringType.getSETTING_ID());
                currentMachineCuringType.setDESCRIPTION(machineCuringType.getDESCRIPTION());
                currentMachineCuringType.setCAVITY(machineCuringType.getCAVITY());
                currentMachineCuringType.setLAST_UPDATE_DATE(new Date());
                currentMachineCuringType.setLAST_UPDATED_BY(machineCuringType.getLAST_UPDATED_BY());
                return machineCuringTypeRepo.save(currentMachineCuringType);
            } else {
                throw new RuntimeException("MachineCuringType with ID " + machineCuringType.getMACHINECURINGTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineCuringType: " + e.getMessage());
            throw e;
        }
    }

    public MachineCuringType deleteMachineCuringType(MachineCuringType machineCuringType) {
        try {
            Optional<MachineCuringType> currentMachineCuringTypeOpt = machineCuringTypeRepo.findById(machineCuringType.getMACHINECURINGTYPE_ID());
            if (currentMachineCuringTypeOpt.isPresent()) {
                MachineCuringType currentMachineCuringType = currentMachineCuringTypeOpt.get();
                currentMachineCuringType.setSTATUS(BigDecimal.valueOf(0));
                currentMachineCuringType.setLAST_UPDATE_DATE(new Date());
                currentMachineCuringType.setLAST_UPDATED_BY(machineCuringType.getLAST_UPDATED_BY());
                return machineCuringTypeRepo.save(currentMachineCuringType);
            } else {
                throw new RuntimeException("MachineCuringType with ID " + machineCuringType.getMACHINECURINGTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineCuringType: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineCuringType restoreMachineCuringType(MachineCuringType machineCuringType) {
        try {
            Optional<MachineCuringType> currentMachineCuringTypeOpt = machineCuringTypeRepo.findById(machineCuringType.getMACHINECURINGTYPE_ID());
            if (currentMachineCuringTypeOpt.isPresent()) {
                MachineCuringType currentMachineCuringType = currentMachineCuringTypeOpt.get();
                currentMachineCuringType.setSTATUS(BigDecimal.valueOf(1)); 
                currentMachineCuringType.setLAST_UPDATE_DATE(new Date());
                currentMachineCuringType.setLAST_UPDATED_BY(machineCuringType.getLAST_UPDATED_BY());
                return machineCuringTypeRepo.save(currentMachineCuringType);
            } else {
                throw new RuntimeException("MachineCuringType with ID " + machineCuringType.getMACHINECURINGTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring machineCuringType: " + e.getMessage());
            throw e;
        }
    }


    public void deleteAllMachineCuringType() {
        machineCuringTypeRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMachineCuringTypesExcel() throws IOException {
        List<MachineCuringType> machineCuringTypes = machineCuringTypeRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineCuringTypes);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineCuringType> machineCuringTypes) throws IOException {
        String[] header = { "NOMOR", "MACHINECURINGTYPE_ID", "SETTING_ID", "DESCRIPTION", "CAVITY" };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("MACHINE CURING TYPE DATA");

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
            int nomor = 1;
            for (MachineCuringType m : machineCuringTypes) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(m.getMACHINECURINGTYPE_ID());
                idCell.setCellStyle(borderStyle);

                Cell settingIdCell = dataRow.createCell(2);
                settingIdCell.setCellValue(m.getSETTING_ID() != null ? m.getSETTING_ID().doubleValue() : null);
                settingIdCell.setCellStyle(borderStyle);

                Cell descriptionCell = dataRow.createCell(3);
                descriptionCell.setCellValue(m.getDESCRIPTION());
                descriptionCell.setCellStyle(borderStyle);

                Cell cavityCell = dataRow.createCell(4);
                cavityCell.setCellValue(m.getCAVITY() != null ? m.getCAVITY().doubleValue() : null);
                cavityCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export Machine Curing Type data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    
    
}