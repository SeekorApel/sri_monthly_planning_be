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
import sri.sysint.sri_starter_back.model.MachineTass;
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.MachineTassRepo;
import sri.sysint.sri_starter_back.repository.MachineTassTypeRepo;

@Service
@Transactional
public class MachineTassServiceImpl {
    @Autowired
    private MachineTassRepo machineTassRepo;
	@Autowired
	private MachineTassTypeRepo machineTassTypeRepo;
    @Autowired
    private BuildingRepo buildingRepo;

    public MachineTassServiceImpl(MachineTassRepo machineTassRepo) {
        this.machineTassRepo = machineTassRepo;
    }

    public List<MachineTass> getAllMachineTass() {
        Iterable<MachineTass> machineTasses = machineTassRepo.getDataOrderId();
        List<MachineTass> machineTassList = new ArrayList<>();
        for (MachineTass machine : machineTasses) {
            MachineTass machineTassTemp = new MachineTass(machine);
            machineTassList.add(machineTassTemp);
        }
        return machineTassList;
    }

    public Optional<MachineTass> getMachineTassById(String id) {
        Optional<MachineTass> machineTass = machineTassRepo.findById(id);
        return machineTass;
    }

    public MachineTass saveMachineTass(MachineTass machineTass) {
        try {
        	machineTass.setID_MACHINE_TASS(machineTass.getID_MACHINE_TASS());
            machineTass.setSTATUS(BigDecimal.valueOf(1));
            machineTass.setCREATION_DATE(new Date());
            machineTass.setLAST_UPDATE_DATE(new Date());
            return machineTassRepo.save(machineTass);
        } catch (Exception e) {
            System.err.println("Error saving machineTass: " + e.getMessage());
            throw e;
        }
    }

    public MachineTass updateMachineTass(MachineTass machineTass) {
        try {
            Optional<MachineTass> currentMachineTassOpt = machineTassRepo.findById(machineTass.getID_MACHINE_TASS());
            if (currentMachineTassOpt.isPresent()) {
                MachineTass currentMachineTass = currentMachineTassOpt.get();
                
                currentMachineTass.setBUILDING_ID(machineTass.getBUILDING_ID());
                currentMachineTass.setFLOOR(machineTass.getFLOOR());
                currentMachineTass.setMACHINE_NUMBER(machineTass.getMACHINE_NUMBER());
                currentMachineTass.setMACHINE_TASS_TYPE_ID(machineTass.getMACHINE_TASS_TYPE_ID());
                currentMachineTass.setWORK_CENTER_TEXT(machineTass.getWORK_CENTER_TEXT());
                
                currentMachineTass.setLAST_UPDATE_DATE(new Date());
                currentMachineTass.setLAST_UPDATED_BY(machineTass.getLAST_UPDATED_BY());
                return machineTassRepo.save(currentMachineTass);
            } else {
                throw new RuntimeException("MachineTass with ID " + machineTass.getID_MACHINE_TASS() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineTass: " + e.getMessage());
            throw e;
        }
    }

    public MachineTass deleteMachineTass(MachineTass machineTass) {
        try {
            Optional<MachineTass> currentMachineTassOpt = machineTassRepo.findById(machineTass.getID_MACHINE_TASS());
            if (currentMachineTassOpt.isPresent()) {
                MachineTass currentMachineTass = currentMachineTassOpt.get();
                currentMachineTass.setSTATUS(BigDecimal.valueOf(0));
                currentMachineTass.setLAST_UPDATE_DATE(new Date());
                currentMachineTass.setLAST_UPDATED_BY(machineTass.getLAST_UPDATED_BY());
                return machineTassRepo.save(currentMachineTass);
            } else {
                throw new RuntimeException("MachineTass with ID " + machineTass.getID_MACHINE_TASS() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineTass: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineTass restoreMachineTass(MachineTass machineTass) {
        try {
            Optional<MachineTass> currentMachineTassOpt = machineTassRepo.findById(machineTass.getID_MACHINE_TASS());
            
            if (currentMachineTassOpt.isPresent()) {
                MachineTass currentMachineTass = currentMachineTassOpt.get();
                
                currentMachineTass.setSTATUS(BigDecimal.valueOf(1)); 
                currentMachineTass.setLAST_UPDATE_DATE(new Date());
                currentMachineTass.setLAST_UPDATED_BY(machineTass.getLAST_UPDATED_BY());
                
                return machineTassRepo.save(currentMachineTass);
            } else {
                throw new RuntimeException("MachineTass with ID " + machineTass.getID_MACHINE_TASS() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring machineTass: " + e.getMessage());
            throw e;
        }
    }


    public void deleteAllMachineTass() {
        machineTassRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMachineTassExcel() throws IOException {
        List<MachineTass> machineTasss = machineTassRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineTasss);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineTass> machineTasses) throws IOException {
        String[] header = {
            "NOMOR",
            "ID_MACHINE_TASS",
            "BUILDING_NAME",
            "FLOOR",
            "MACHINE_NUMBER",
            "MACHINE_TASS_TYPE",
            "WORK_CENTER_TEXT"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<MachineTassType> activeMachineTassTypes = machineTassTypeRepo.findMachineTassTypeActive();

            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            List<String> machineTassTypes = activeMachineTassTypes.stream()
                .map(MachineTassType::getMACHINETASSTYPE_ID)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE TASS DATA");
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

            Sheet hiddenBuildingSheet = workbook.createSheet("HIDDEN_BUILDINGS");
            for (int i = 0; i < buildingNames.size(); i++) {
                Row row = hiddenBuildingSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(buildingNames.get(i));
            }

            Name buildingNameRange = workbook.createName();
            buildingNameRange.setNameName("BuildingNames");
            buildingNameRange.setRefersToFormula("HIDDEN_BUILDINGS!$A$1:$A$" + buildingNames.size());
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenBuildingSheet), true);

            Sheet hiddenMachineTassSheet = workbook.createSheet("HIDDEN_TASS_TYPES");
            for (int i = 0; i < machineTassTypes.size(); i++) {
                Row row = hiddenMachineTassSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineTassTypes.get(i));
            }

            Name tassTypeRange = workbook.createName();
            tassTypeRange.setNameName("MachineTassTypes");
            tassTypeRange.setRefersToFormula("HIDDEN_TASS_TYPES!$A$1:$A$" + machineTassTypes.size());
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenMachineTassSheet), true);

            int rowIndex = 1;
            for (MachineTass m : machineTasses) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(m.getID_MACHINE_TASS());
                idCell.setCellStyle(borderStyle);

                String buildingName = null;
                if (m.getBUILDING_ID() != null) {
                    Building building = buildingRepo.findById(m.getBUILDING_ID()).orElse(null);
                    if (building != null) {
                        buildingName = building.getBUILDING_NAME();
                    }
                }
                Cell buildingNameCell = dataRow.createCell(2);
                buildingNameCell.setCellValue(buildingName != null ? buildingName : "");
                buildingNameCell.setCellStyle(borderStyle);

                Cell floorCell = dataRow.createCell(3);
                floorCell.setCellValue(m.getFLOOR() != null ? m.getFLOOR().doubleValue() : null);
                floorCell.setCellStyle(borderStyle);

                Cell machineNumberCell = dataRow.createCell(4);
                machineNumberCell.setCellValue(m.getMACHINE_NUMBER() != null ? m.getMACHINE_NUMBER().doubleValue() : null);
                machineNumberCell.setCellStyle(borderStyle);

                String typeName = null;
                if (m.getMACHINE_TASS_TYPE_ID() != null) {
                    MachineTassType type = machineTassTypeRepo.findById(m.getMACHINE_TASS_TYPE_ID()).orElse(null);
                    if (type != null) {
                    	typeName = type.getMACHINETASSTYPE_ID();
                    }
                }
                
                Cell typeCell = dataRow.createCell(5);
                typeCell.setCellValue(typeName != null ? typeName : "");
                typeCell.setCellStyle(borderStyle);

                Cell workCenterTextCell = dataRow.createCell(6);
                workCenterTextCell.setCellValue(m.getWORK_CENTER_TEXT() != null ? m.getWORK_CENTER_TEXT() : "");
                workCenterTextCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint buildingConstraint = validationHelper.createFormulaListConstraint("BuildingNames");
            CellRangeAddressList buildingAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation buildingValidation = validationHelper.createValidation(buildingConstraint, buildingAddressList);
            buildingValidation.setSuppressDropDownArrow(true);
            buildingValidation.setShowErrorBox(true);
            sheet.addValidationData(buildingValidation);

            DataValidationConstraint tassTypeConstraint = validationHelper.createFormulaListConstraint("MachineTassTypes");
            CellRangeAddressList tassTypeAddressList = new CellRangeAddressList(1, 1000, 5, 5);
            DataValidation tassTypeValidation = validationHelper.createValidation(tassTypeConstraint, tassTypeAddressList);
            tassTypeValidation.setSuppressDropDownArrow(true);
            tassTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(tassTypeValidation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data MachineTass");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }


    public ByteArrayInputStream layoutMachineTassExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "ID_MACHINE_TASS",
            "BUILDING_NAME",
            "FLOOR",
            "MACHINE_NUMBER",
            "MACHINE_TASS_TYPE",
            "WORK_CENTER_TEXT"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<MachineTassType> activeMachineTassTypes = machineTassTypeRepo.findMachineTassTypeActive();

            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            List<String> machineTassTypes = activeMachineTassTypes.stream()
                .map(MachineTassType::getMACHINETASSTYPE_ID)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE TASS DATA");
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

            Sheet hiddenBuildingSheet = workbook.createSheet("HIDDEN_BUILDINGS");
            for (int i = 0; i < buildingNames.size(); i++) {
                Row row = hiddenBuildingSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(buildingNames.get(i));
            }

            Name buildingNameRange = workbook.createName();
            buildingNameRange.setNameName("BuildingNames");
            buildingNameRange.setRefersToFormula("HIDDEN_BUILDINGS!$A$1:$A$" + buildingNames.size());
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenBuildingSheet), true);

            Sheet hiddenMachineTassSheet = workbook.createSheet("HIDDEN_TASS_TYPES");
            for (int i = 0; i < machineTassTypes.size(); i++) {
                Row row = hiddenMachineTassSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineTassTypes.get(i));
            }

            Name tassTypeRange = workbook.createName();
            tassTypeRange.setNameName("MachineTassTypes");
            tassTypeRange.setRefersToFormula("HIDDEN_TASS_TYPES!$A$1:$A$" + machineTassTypes.size());
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenMachineTassSheet), true);

            int rowIndex = 1;
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint buildingConstraint = validationHelper.createFormulaListConstraint("BuildingNames");
            CellRangeAddressList buildingAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation buildingValidation = validationHelper.createValidation(buildingConstraint, buildingAddressList);
            buildingValidation.setSuppressDropDownArrow(true);
            buildingValidation.setShowErrorBox(true);
            sheet.addValidationData(buildingValidation);

            DataValidationConstraint tassTypeConstraint = validationHelper.createFormulaListConstraint("MachineTassTypes");
            CellRangeAddressList tassTypeAddressList = new CellRangeAddressList(1, 1000, 5, 5);
            DataValidation tassTypeValidation = validationHelper.createValidation(tassTypeConstraint, tassTypeAddressList);
            tassTypeValidation.setSuppressDropDownArrow(true);
            tassTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(tassTypeValidation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data MachineTass");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
