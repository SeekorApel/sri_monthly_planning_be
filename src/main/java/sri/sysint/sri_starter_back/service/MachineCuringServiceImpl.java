
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
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;


@Service
@Transactional
public class MachineCuringServiceImpl {
    @Autowired
    private MachineCuringRepo machineCuringRepo;
    @Autowired
    private MachineCuringTypeRepo machineCuringTypeRepo;
    @Autowired
    private BuildingRepo buildingRepo;

    public MachineCuringServiceImpl(MachineCuringRepo machineCuringRepo) {
        this.machineCuringRepo = machineCuringRepo;
    }

    public List<MachineCuring> getAllMachineCuring() {
        Iterable<MachineCuring> machineCurings = machineCuringRepo.getDataOrderId();
        List<MachineCuring> machineCuringList = new ArrayList<>();
        for (MachineCuring machine : machineCurings) {
            MachineCuring machineCuringTemp = new MachineCuring(machine);
            machineCuringList.add(machineCuringTemp);
        }
        return machineCuringList;
    }

    public Optional<MachineCuring> getMachineCuringById(String id) {
        Optional<MachineCuring> machineCuring = machineCuringRepo.findById(id);
        return machineCuring;
    }

    public MachineCuring saveMachineCuring(MachineCuring machineCuring) {
        try {
            machineCuring.setWORK_CENTER_TEXT(machineCuring.getWORK_CENTER_TEXT());
            machineCuring.setSTATUS(BigDecimal.valueOf(1));
            machineCuring.setCREATION_DATE(new Date());
            machineCuring.setLAST_UPDATE_DATE(new Date());
            return machineCuringRepo.save(machineCuring);
        } catch (Exception e) {
            System.err.println("Error saving machineCuring: " + e.getMessage());
            throw e;
        }
    }

    public MachineCuring updateMachineCuring(MachineCuring machineCuring) {
        try {
            Optional<MachineCuring> currentMachineCuringOpt = machineCuringRepo.findById(machineCuring.getWORK_CENTER_TEXT());
            if (currentMachineCuringOpt.isPresent()) {
                MachineCuring currentMachineCuring = currentMachineCuringOpt.get();
                currentMachineCuring.setLAST_UPDATE_DATE(new Date());
                currentMachineCuring.setBUILDING_ID(machineCuring.getBUILDING_ID());
                currentMachineCuring.setCAVITY(machineCuring.getCAVITY());
                currentMachineCuring.setMACHINE_TYPE(machineCuring.getMACHINE_TYPE());
                currentMachineCuring.setSTATUS_USAGE(machineCuring.getSTATUS_USAGE());


                currentMachineCuring.setLAST_UPDATED_BY(machineCuring.getLAST_UPDATED_BY());
                return machineCuringRepo.save(currentMachineCuring);
            } else {
                throw new RuntimeException("MachineCuring with ID " + machineCuring.getWORK_CENTER_TEXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineCuring: " + e.getMessage());
            throw e;
        }
    }

    public MachineCuring deleteMachineCuring(MachineCuring machineCuring) {
        try {
            Optional<MachineCuring> currentMachineCuringOpt = machineCuringRepo.findById(machineCuring.getWORK_CENTER_TEXT());
            if (currentMachineCuringOpt.isPresent()) {
                MachineCuring currentMachineCuring = currentMachineCuringOpt.get();
                currentMachineCuring.setSTATUS(BigDecimal.valueOf(0));
                currentMachineCuring.setLAST_UPDATE_DATE(new Date());
                currentMachineCuring.setLAST_UPDATED_BY(machineCuring.getLAST_UPDATED_BY());
                return machineCuringRepo.save(currentMachineCuring);
            } else {
                throw new RuntimeException("MachineCuring with ID " + machineCuring.getWORK_CENTER_TEXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineCuring: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineCuring restoreMachineCuring(MachineCuring machineCuring) {
        try {
            Optional<MachineCuring> currentMachineCuringOpt = machineCuringRepo.findById(machineCuring.getWORK_CENTER_TEXT());
            
            if (currentMachineCuringOpt.isPresent()) {
                MachineCuring currentMachineCuring = currentMachineCuringOpt.get();
                
                currentMachineCuring.setSTATUS(BigDecimal.valueOf(1)); 
                currentMachineCuring.setLAST_UPDATE_DATE(new Date());
                currentMachineCuring.setLAST_UPDATED_BY(machineCuring.getLAST_UPDATED_BY());
                
                return machineCuringRepo.save(currentMachineCuring);
            } else {
                throw new RuntimeException("MachineCuring with ID " + machineCuring.getWORK_CENTER_TEXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring machineCuring: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllMachineCuring() {
        machineCuringRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMachineCuringsExcel() throws IOException {
        List<MachineCuring> machineCurings = machineCuringRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineCurings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineCuring> machineCurings) throws IOException {
        String[] header = {
            "NOMOR",
            "WORK_CENTER_TEXT",
            "BUILDING_NAME",
            "CAVITY",
            "MACHINE_TYPE",
            "STATUS_USAGE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            List<MachineCuringType> activeMachineTypes = machineCuringTypeRepo.findMachineCuringTypeActive();
            List<String> machineTypes = activeMachineTypes.stream()
                .map(MachineCuringType::getMACHINECURINGTYPE_ID)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE CURING DATA");
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

            Name buildingRange = workbook.createName();
            buildingRange.setNameName("BuildingNames");
            buildingRange.setRefersToFormula("HIDDEN_BUILDINGS!$A$1:$A$" + buildingNames.size());

            Sheet hiddenMachineTypeSheet = workbook.createSheet("HIDDEN_MACHINE_TYPES");
            for (int i = 0; i < machineTypes.size(); i++) {
                Row row = hiddenMachineTypeSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineTypes.get(i));
            }

            Name machineTypeRange = workbook.createName();
            machineTypeRange.setNameName("MachineTypes");
            machineTypeRange.setRefersToFormula("HIDDEN_MACHINE_TYPES!$A$1:$A$" + machineTypes.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenBuildingSheet), true);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenMachineTypeSheet), true);

            int rowIndex = 1;
            int nomor = 1;
            for (MachineCuring mc : machineCurings) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell workcentertextCell = dataRow.createCell(1);
                workcentertextCell.setCellValue(mc.getWORK_CENTER_TEXT());
                workcentertextCell.setCellStyle(borderStyle);

                String buildingName = null;
                if (mc.getBUILDING_ID() != null) {
                    Building building = buildingRepo.findById(mc.getBUILDING_ID()).orElse(null);
                    if (building != null) {
                        buildingName = building.getBUILDING_NAME();
                    }
                }
                
                Cell buildingNameCell = dataRow.createCell(2);
                buildingNameCell.setCellValue(buildingName != null ? buildingName : "");
                buildingNameCell.setCellStyle(borderStyle);

                Cell cavityCell = dataRow.createCell(3);
                cavityCell.setCellValue(mc.getCAVITY().doubleValue());
                cavityCell.setCellStyle(borderStyle);

                String typeName = null;
                if (mc.getMACHINE_TYPE() != null) {
                    MachineCuringType type = machineCuringTypeRepo.findById(mc.getMACHINE_TYPE()).orElse(null);
                    if (type != null) {
                    	typeName = type.getMACHINECURINGTYPE_ID();
                    }
                }
                
                Cell machineTypeCell = dataRow.createCell(4);
                machineTypeCell.setCellValue(typeName != null ? typeName : "");
                machineTypeCell.setCellStyle(borderStyle);

                Cell statusUsageCell = dataRow.createCell(5);
                statusUsageCell.setCellValue(mc.getSTATUS_USAGE() != null ? mc.getSTATUS_USAGE().doubleValue() : 0.0);
                statusUsageCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint buildingConstraint = validationHelper.createFormulaListConstraint("BuildingNames");
            CellRangeAddressList buildingAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation buildingValidation = validationHelper.createValidation(buildingConstraint, buildingAddressList);
            buildingValidation.setSuppressDropDownArrow(true);
            buildingValidation.setShowErrorBox(true);
            sheet.addValidationData(buildingValidation);

            DataValidationConstraint machineTypeConstraint = validationHelper.createFormulaListConstraint("MachineTypes");
            CellRangeAddressList machineTypeAddressList = new CellRangeAddressList(1, 1000, 4, 4);
            DataValidation machineTypeValidation = validationHelper.createValidation(machineTypeConstraint, machineTypeAddressList);
            machineTypeValidation.setSuppressDropDownArrow(true);
            machineTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(machineTypeValidation);

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            throw new IOException("Failed to export MachineCuring data", e);
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutMachineCuringsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "WORK_CENTER_TEXT",
            "BUILDING_NAME",
            "CAVITY",
            "MACHINE_TYPE",
            "STATUS_USAGE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            List<MachineCuringType> activeMachineTypes = machineCuringTypeRepo.findMachineCuringTypeActive();
            List<String> machineTypes = activeMachineTypes.stream()
                .map(MachineCuringType::getMACHINECURINGTYPE_ID)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE CURING DATA");
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

            Name buildingRange = workbook.createName();
            buildingRange.setNameName("BuildingNames");
            buildingRange.setRefersToFormula("HIDDEN_BUILDINGS!$A$1:$A$" + buildingNames.size());

            Sheet hiddenMachineTypeSheet = workbook.createSheet("HIDDEN_MACHINE_TYPES");
            for (int i = 0; i < machineTypes.size(); i++) {
                Row row = hiddenMachineTypeSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineTypes.get(i));
            }

            Name machineTypeRange = workbook.createName();
            machineTypeRange.setNameName("MachineTypes");
            machineTypeRange.setRefersToFormula("HIDDEN_MACHINE_TYPES!$A$1:$A$" + machineTypes.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenBuildingSheet), true);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenMachineTypeSheet), true);

            int rowIndex = 1;
            int nomor = 1;
            
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

            DataValidationConstraint machineTypeConstraint = validationHelper.createFormulaListConstraint("MachineTypes");
            CellRangeAddressList machineTypeAddressList = new CellRangeAddressList(1, 1000, 4, 4);
            DataValidation machineTypeValidation = validationHelper.createValidation(machineTypeConstraint, machineTypeAddressList);
            machineTypeValidation.setSuppressDropDownArrow(true);
            machineTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(machineTypeValidation);

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            throw new IOException("Failed to export MachineCuring data", e);
        } finally {
            workbook.close();
            out.close();
        }
    }
}
