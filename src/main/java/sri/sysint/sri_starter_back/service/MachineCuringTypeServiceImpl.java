package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.Setting;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;
import sri.sysint.sri_starter_back.repository.SettingRepo;


@Service
@Transactional
public class MachineCuringTypeServiceImpl {
	@Autowired
    private MachineCuringTypeRepo machineCuringTypeRepo;
	@Autowired
    private SettingRepo settingRepo;
	
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
        String[] header = { "NOMOR", "MACHINECURINGTYPE_ID", "SETTING_VALUE", "DESCRIPTION", "CAVITY" };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Setting> activeSettings = settingRepo.findSettingActive();
            List<String> settingValues = activeSettings.stream()
                .map(Setting::getSETTING_VALUE)
                .collect(Collectors.toList());

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
            borderStyle.setAlignment(HorizontalAlignment.CENTER);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_SETTINGS");
            for (int i = 0; i < settingValues.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(settingValues.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("SettingValues");
            namedRange.setRefersToFormula("HIDDEN_SETTINGS!$A$1:$A$" + settingValues.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

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

                Cell settingValueCell = dataRow.createCell(2);
                settingValueCell.setCellStyle(borderStyle);
                
                String settingValue = "";
                if (m.getSETTING_ID() != null) {
                    Optional<Setting> settingOpt = settingRepo.findById(m.getSETTING_ID());
                    if (settingOpt.isPresent()) {
                        settingValue = settingOpt.get().getSETTING_VALUE();
                    }
                }
                settingValueCell.setCellValue(settingValue);

                Cell descriptionCell = dataRow.createCell(3);
                descriptionCell.setCellValue(m.getDESCRIPTION());
                descriptionCell.setCellStyle(borderStyle);

                Cell cavityCell = dataRow.createCell(4);
                cavityCell.setCellValue(m.getCAVITY() != null ? m.getCAVITY().doubleValue() : null);
                cavityCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("SettingValues");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

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

    public ByteArrayInputStream layoutMachineCuringTypesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = { "NOMOR", "MACHINECURINGTYPE_ID", "SETTING_VALUE", "DESCRIPTION", "CAVITY" };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Setting> activeSettings = settingRepo.findSettingActive();
            List<String> settingValues = activeSettings.stream()
                .map(Setting::getSETTING_VALUE)
                .collect(Collectors.toList());

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
            borderStyle.setAlignment(HorizontalAlignment.CENTER);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_SETTINGS");
            for (int i = 0; i < settingValues.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(settingValues.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("SettingValues");
            namedRange.setRefersToFormula("HIDDEN_SETTINGS!$A$1:$A$" + settingValues.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            int nomor = 1;
            

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("SettingValues");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

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