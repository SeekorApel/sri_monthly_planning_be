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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;


import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineExtruding;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.MachineExtrudingRepo;

@Service
@Transactional
public class MachineExtrudingServiceImpl {
    @Autowired
    private MachineExtrudingRepo machineExtrudingRepo;
    
    @Autowired
    private BuildingRepo buildingRepo;
    
    public MachineExtrudingServiceImpl(MachineExtrudingRepo machineExtrudingRepo){
        this.machineExtrudingRepo = machineExtrudingRepo;
    }
    
    public BigDecimal getNewId() {
        return machineExtrudingRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<MachineExtruding> getAllMachineExtruding() {
        Iterable<MachineExtruding> machineExtrudings = machineExtrudingRepo.getDataOrderId();
        List<MachineExtruding> machineExtrudingList = new ArrayList<>();
        for (MachineExtruding item : machineExtrudings) {
            MachineExtruding machineExtrudingTemp = new MachineExtruding(item);
            machineExtrudingList.add(machineExtrudingTemp);
        }
        
        return machineExtrudingList;
    }
    
    public Optional<MachineExtruding> getMachineExtrudingById(BigDecimal id) {
        Optional<MachineExtruding> machineExtruding = machineExtrudingRepo.findById(id);
        return machineExtruding;
    }
    
    public MachineExtruding saveMachineExtruding(MachineExtruding machineExtruding) {
        try {
            machineExtruding.setID_MACHINE_EXT(getNewId());
            machineExtruding.setSTATUS(BigDecimal.valueOf(1));
            machineExtruding.setCREATION_DATE(new Date());
            machineExtruding.setLAST_UPDATE_DATE(new Date());
            return machineExtrudingRepo.save(machineExtruding);
        } catch (Exception e) {
            System.err.println("Error saving machineExtruding: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineExtruding updateMachineExtruding(MachineExtruding machineExtruding) {
        try {
            Optional<MachineExtruding> currentMachineExtrudingOpt = machineExtrudingRepo.findById(machineExtruding.getID_MACHINE_EXT());
            
            if (currentMachineExtrudingOpt.isPresent()) {
                MachineExtruding currentMachineExtruding = currentMachineExtrudingOpt.get();
                
                currentMachineExtruding.setBUILDING_ID(machineExtruding.getBUILDING_ID());
                currentMachineExtruding.setTYPE(machineExtruding.getTYPE());
                currentMachineExtruding.setLAST_UPDATE_DATE(new Date());
                currentMachineExtruding.setLAST_UPDATED_BY(machineExtruding.getLAST_UPDATED_BY());
                
                return machineExtrudingRepo.save(currentMachineExtruding);
            } else {
                throw new RuntimeException("MachineExtruding with ID " + machineExtruding.getID_MACHINE_EXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineExtruding: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineExtruding deleteMachineExtruding(MachineExtruding machineExtruding) {
        try {
            Optional<MachineExtruding> currentMachineExtrudingOpt = machineExtrudingRepo.findById(machineExtruding.getID_MACHINE_EXT());
            
            if (currentMachineExtrudingOpt.isPresent()) {
                MachineExtruding currentMachineExtruding = currentMachineExtrudingOpt.get();
                
                currentMachineExtruding.setSTATUS(BigDecimal.valueOf(0));
                currentMachineExtruding.setLAST_UPDATE_DATE(new Date());
                currentMachineExtruding.setLAST_UPDATED_BY(machineExtruding.getLAST_UPDATED_BY());
                
                return machineExtrudingRepo.save(currentMachineExtruding);
            } else {
                throw new RuntimeException("MachineExtruding with ID " + machineExtruding.getID_MACHINE_EXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineExtruding: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineExtruding restoreMachineExtruding(MachineExtruding machineExtruding) {
        try {
            Optional<MachineExtruding> currentMachineExtrudingOpt = machineExtrudingRepo.findById(machineExtruding.getID_MACHINE_EXT());
            
            if (currentMachineExtrudingOpt.isPresent()) {
                MachineExtruding currentMachineExtruding = currentMachineExtrudingOpt.get();
                
                currentMachineExtruding.setSTATUS(BigDecimal.valueOf(1));
                currentMachineExtruding.setLAST_UPDATE_DATE(new Date());
                currentMachineExtruding.setLAST_UPDATED_BY(machineExtruding.getLAST_UPDATED_BY());
                
                return machineExtrudingRepo.save(currentMachineExtruding);
            } else {
                throw new RuntimeException("MachineExtruding with ID " + machineExtruding.getID_MACHINE_EXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring machineExtruding: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllMachineExtruding() {
        machineExtrudingRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMachineExtrudingsExcel() throws IOException {
        List<MachineExtruding> machineExtrudings = machineExtrudingRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineExtrudings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineExtruding> machineExtrudings) throws IOException {
        String[] header = {
            "NOMOR",
            "ID_MACHINE_EXT",
            "BUILDING_NAME",
            "TYPE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE EXTRUDING DATA");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_BUILDINGS");
            for (int i = 0; i < buildingNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(buildingNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("BuildingNames");
            namedRange.setRefersToFormula("HIDDEN_BUILDINGS!$A$1:$A$" + buildingNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            for (MachineExtruding m : machineExtrudings) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(m.getID_MACHINE_EXT() != null ? m.getID_MACHINE_EXT().doubleValue() : null);
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

                Cell typeCell = dataRow.createCell(3);
                typeCell.setCellValue(m.getTYPE() != null ? m.getTYPE() : "");
                typeCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("BuildingNames");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Machine Extruding");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutMachineExtrudingsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "ID_MACHINE_EXT",
            "BUILDING_NAME",
            "TYPE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE EXTRUDING DATA");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_BUILDINGS");
            for (int i = 0; i < buildingNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(buildingNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("BuildingNames");
            namedRange.setRefersToFormula("HIDDEN_BUILDINGS!$A$1:$A$" + buildingNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("BuildingNames");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Machine Extruding");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
}
