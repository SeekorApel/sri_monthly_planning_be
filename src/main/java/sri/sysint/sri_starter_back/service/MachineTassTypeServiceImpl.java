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
import org.apache.poi.ss.usermodel.DataValidationHelper;
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
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.model.Setting;
import sri.sysint.sri_starter_back.repository.MachineTassTypeRepo;
import sri.sysint.sri_starter_back.repository.SettingRepo;


@Service
@Transactional
public class MachineTassTypeServiceImpl {
	@Autowired
	private MachineTassTypeRepo machineTassTypeRepo;
	@Autowired
	private SettingRepo settingRepo;
	public MachineTassTypeServiceImpl(MachineTassTypeRepo machineTassTypeRepo) {
        this.machineTassTypeRepo = machineTassTypeRepo;
    }

    public List<MachineTassType> getAllMachineTassType() {
        Iterable<MachineTassType> machineTassTypes = machineTassTypeRepo.getDataOrderId();
        List<MachineTassType> machineTassTypeList = new ArrayList<>();
        for (MachineTassType item : machineTassTypes) {
            MachineTassType machineTassTypeTemp = new MachineTassType(item);
            machineTassTypeList.add(machineTassTypeTemp);
        }
        return machineTassTypeList;
    }

    public Optional<MachineTassType> getMachineTassTypeById(String id) {
        Optional<MachineTassType> machineTassType = machineTassTypeRepo.findById(id);
        return machineTassType;
    }

    public MachineTassType saveMachineTassType(MachineTassType machineTassType) {
        try {
        	machineTassType.setMACHINETASSTYPE_ID(machineTassType.getMACHINETASSTYPE_ID());
            machineTassType.setSTATUS(BigDecimal.valueOf(1));
            machineTassType.setCREATION_DATE(new Date());
            machineTassType.setLAST_UPDATE_DATE(new Date());
            return machineTassTypeRepo.save(machineTassType);
        } catch (Exception e) {
            System.err.println("Error saving machineTassType: " + e.getMessage());
            throw e;
        }
    }

    public MachineTassType updateMachineTassType(MachineTassType machineTassType) {
        try {
            Optional<MachineTassType> currentMachineTassTypeOpt = machineTassTypeRepo.findById(machineTassType.getMACHINETASSTYPE_ID());
            if (currentMachineTassTypeOpt.isPresent()) {
                MachineTassType currentMachineTassType = currentMachineTassTypeOpt.get();
                currentMachineTassType.setSETTING_ID(machineTassType.getSETTING_ID());
                currentMachineTassType.setDESCRIPTION(machineTassType.getDESCRIPTION());
                currentMachineTassType.setLAST_UPDATE_DATE(new Date());
                currentMachineTassType.setLAST_UPDATED_BY(machineTassType.getLAST_UPDATED_BY());
                return machineTassTypeRepo.save(currentMachineTassType);
            } else {
                throw new RuntimeException("MachineTassType with ID " + machineTassType.getMACHINETASSTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineTassType: " + e.getMessage());
            throw e;
        }
    }

    public MachineTassType deleteMachineTassType(MachineTassType machineTassType) {
        try {
            Optional<MachineTassType> currentMachineTassTypeOpt = machineTassTypeRepo.findById(machineTassType.getMACHINETASSTYPE_ID());
            if (currentMachineTassTypeOpt.isPresent()) {
                MachineTassType currentMachineTassType = currentMachineTassTypeOpt.get();
                currentMachineTassType.setSTATUS(BigDecimal.valueOf(0));
                currentMachineTassType.setLAST_UPDATE_DATE(new Date());
                currentMachineTassType.setLAST_UPDATED_BY(machineTassType.getLAST_UPDATED_BY());
                return machineTassTypeRepo.save(currentMachineTassType);
            } else {
                throw new RuntimeException("MachineTassType with ID " + machineTassType.getMACHINETASSTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineTassType: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineTassType restoreMachineTassType(MachineTassType machineTassType) {
        try {
            Optional<MachineTassType> currentMachineTassTypeOpt = machineTassTypeRepo.findById(machineTassType.getMACHINETASSTYPE_ID());
            
            if (currentMachineTassTypeOpt.isPresent()) {
                MachineTassType currentMachineTassType = currentMachineTassTypeOpt.get();
                
                currentMachineTassType.setSTATUS(BigDecimal.valueOf(1)); 
                currentMachineTassType.setLAST_UPDATE_DATE(new Date());
                currentMachineTassType.setLAST_UPDATED_BY(machineTassType.getLAST_UPDATED_BY());
                
                return machineTassTypeRepo.save(currentMachineTassType);
            } else {
                throw new RuntimeException("MachineTassType with ID " + machineTassType.getMACHINETASSTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring machineTassType: " + e.getMessage());
            throw e;
        }
    }


    public void deleteAllMachineTassType() {
        machineTassTypeRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMachineTassTypesExcel() throws IOException {
        List<MachineTassType> machineTassTypes = machineTassTypeRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineTassTypes);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineTassType> machineTassTypes) throws IOException {
        String[] header = { "NOMOR", "MACHINETASSTYPE_ID", "SETTING_VALUE", "DESCRIPTION" };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Setting> activeSettings = settingRepo.findSettingActive();
            List<String> settingValues = activeSettings.stream()
                .map(Setting::getSETTING_VALUE)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE_TASS_TYPE DATA");
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
            for (MachineTassType mt : machineTassTypes) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(mt.getMACHINETASSTYPE_ID());
                idCell.setCellStyle(borderStyle);

                Cell settingDescriptionCell = dataRow.createCell(2);
                settingDescriptionCell.setCellStyle(borderStyle);
                if (mt.getSETTING_ID() != null) {
                    Optional<Setting> settingOptional = settingRepo.findById(mt.getSETTING_ID());
                    if (settingOptional.isPresent()) {
                        settingDescriptionCell.setCellValue(settingOptional.get().getSETTING_VALUE());
                    } else {
                        settingDescriptionCell.setCellValue("Unknown");
                    }
                } else {
                    settingDescriptionCell.setCellValue("N/A");
                }

                Cell descriptionCell = dataRow.createCell(3);
                descriptionCell.setCellValue(mt.getDESCRIPTION());
                descriptionCell.setCellStyle(borderStyle);
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
            System.err.println("Failed to export MACHINE_TASS_TYPE data.");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutMachineTassTypesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = { "NOMOR", "MACHINETASSTYPE_ID", "SETTING_VALUE", "DESCRIPTION" };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Setting> activeSettings = settingRepo.findSettingActive();
            List<String> settingValues = activeSettings.stream()
                .map(Setting::getSETTING_VALUE)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE_TASS_TYPE DATA");
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
            System.err.println("Failed to export MACHINE_TASS_TYPE data.");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
}