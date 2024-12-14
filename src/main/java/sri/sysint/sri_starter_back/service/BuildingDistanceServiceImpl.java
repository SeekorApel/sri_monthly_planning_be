package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.repository.BuildingDistanceRepo;
import sri.sysint.sri_starter_back.repository.BuildingRepo;

@Service
@Transactional
public class BuildingDistanceServiceImpl {
	@Autowired
    private BuildingDistanceRepo buildingDistanceRepo;
	@Autowired
    private BuildingRepo buildingRepo;
	
    public BuildingDistanceServiceImpl(BuildingDistanceRepo buildingDistanceRepo){
        this.buildingDistanceRepo = buildingDistanceRepo;
    }
    
    public BigDecimal getNewId() {
    	return buildingDistanceRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<BuildingDistance> getAllBuildingDistance() {
    	Iterable<BuildingDistance> buildingDistances = buildingDistanceRepo.getDataOrderId();
        List<BuildingDistance> buildingDistanceList = new ArrayList<>();
        for (BuildingDistance item : buildingDistances) {
            BuildingDistance buildingDistanceTemp = new BuildingDistance(item);
            buildingDistanceList.add(buildingDistanceTemp);
        }
        
        return buildingDistanceList;
    }
    
    public Optional<BuildingDistance> getBuildingDistanceById(BigDecimal id) {
    	Optional<BuildingDistance> buildingDistance = buildingDistanceRepo.findById(id);
    	return buildingDistance;
    }
    
    public BuildingDistance saveBuildingDistance(BuildingDistance buildingDistance) {
        try {
        	buildingDistance.setID_B_DISTANCE(getNewId());
            buildingDistance.setSTATUS(BigDecimal.valueOf(1));
            buildingDistance.setCREATION_DATE(new Date());
            buildingDistance.setLAST_UPDATE_DATE(new Date());
            return buildingDistanceRepo.save(buildingDistance);
        } catch (Exception e) {
            System.err.println("Error saving buildingDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public BuildingDistance updateBuildingDistance(BuildingDistance buildingDistance) {
        try {
            Optional<BuildingDistance> currentBuildingDistanceOpt = buildingDistanceRepo.findById(buildingDistance.getID_B_DISTANCE());
            
            if (currentBuildingDistanceOpt.isPresent()) {
                BuildingDistance currentBuildingDistance = currentBuildingDistanceOpt.get();
                
                currentBuildingDistance.setBUILDING_ID_1(buildingDistance.getBUILDING_ID_1());;
                currentBuildingDistance.setBUILDING_ID_2(buildingDistance.getBUILDING_ID_2());;
                currentBuildingDistance.setDISTANCE(buildingDistance.getDISTANCE());
                currentBuildingDistance.setLAST_UPDATE_DATE(new Date());
                currentBuildingDistance.setLAST_UPDATED_BY(buildingDistance.getLAST_UPDATED_BY());
                
                return buildingDistanceRepo.save(currentBuildingDistance);
            } else {
                throw new RuntimeException("BuildingDistance with ID " + buildingDistance.getID_B_DISTANCE() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating buildingDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public BuildingDistance deleteBuildingDistance(BuildingDistance buildingDistance) {
        try {
            Optional<BuildingDistance> currentBuildingDistanceOpt = buildingDistanceRepo.findById(buildingDistance.getID_B_DISTANCE());
            
            if (currentBuildingDistanceOpt.isPresent()) {
                BuildingDistance currentBuildingDistance = currentBuildingDistanceOpt.get();
                
                currentBuildingDistance.setSTATUS(BigDecimal.valueOf(0));
                currentBuildingDistance.setLAST_UPDATE_DATE(new Date());
                currentBuildingDistance.setLAST_UPDATED_BY(buildingDistance.getLAST_UPDATED_BY());
                
                return buildingDistanceRepo.save(currentBuildingDistance);
            } else {
                throw new RuntimeException("BuildingDistance with ID " + buildingDistance.getID_B_DISTANCE() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating buildingDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public BuildingDistance restoreBuildingDistance(BuildingDistance buildingDistance) {
        try {
            Optional<BuildingDistance> currentBuildingDistanceOpt = buildingDistanceRepo.findById(buildingDistance.getID_B_DISTANCE());
            
            if (currentBuildingDistanceOpt.isPresent()) {
                BuildingDistance currentBuildingDistance = currentBuildingDistanceOpt.get();
                
                currentBuildingDistance.setSTATUS(BigDecimal.valueOf(1)); 
                currentBuildingDistance.setLAST_UPDATE_DATE(new Date());
                currentBuildingDistance.setLAST_UPDATED_BY(buildingDistance.getLAST_UPDATED_BY());
                
                return buildingDistanceRepo.save(currentBuildingDistance);
            } else {
                throw new RuntimeException("BuildingDistance with ID " + buildingDistance.getID_B_DISTANCE() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring buildingDistance: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllBuildingDistance() {
    	buildingDistanceRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportBuildingDistancesExcel() throws IOException {
        List<BuildingDistance> buildingDistances  = buildingDistanceRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(buildingDistances);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<BuildingDistance> buildingDistances) throws IOException {
        String[] header = {
            "NOMOR",
            "ID_B_DISTANCE",
            "BUILDING_NAME_1",
            "BUILDING_NAME_2",
            "DISTANCE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("BUILDING_DISTANCE DATA");
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
            int nomor = 1;
            for (BuildingDistance bd : buildingDistances) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(bd.getID_B_DISTANCE().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell buildingName1Cell = dataRow.createCell(2);
                buildingName1Cell.setCellStyle(borderStyle);

                String buildingName1 = "";
                if (bd.getBUILDING_ID_1() != null) {
                    Optional<Building> building1Opt = activeBuildings.stream()
                        .filter(b -> b.getBUILDING_ID().equals(bd.getBUILDING_ID_1()))
                        .findFirst();
                    buildingName1 = building1Opt.map(Building::getBUILDING_NAME).orElse("");
                }
                buildingName1Cell.setCellValue(buildingName1);

                Cell buildingName2Cell = dataRow.createCell(3);
                buildingName2Cell.setCellStyle(borderStyle);

                String buildingName2 = "";
                if (bd.getBUILDING_ID_2() != null) {
                    Optional<Building> building2Opt = activeBuildings.stream()
                        .filter(b -> b.getBUILDING_ID().equals(bd.getBUILDING_ID_2()))
                        .findFirst();
                    buildingName2 = building2Opt.map(Building::getBUILDING_NAME).orElse("");
                }
                buildingName2Cell.setCellValue(buildingName2);

                Cell distanceCell = dataRow.createCell(4);
                distanceCell.setCellValue(bd.getDISTANCE() != null ? bd.getDISTANCE().doubleValue() : null);
                distanceCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("BuildingNames");
            CellRangeAddressList addressList1 = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation1 = validationHelper.createValidation(constraint, addressList1);
            validation1.setSuppressDropDownArrow(true);
            validation1.setShowErrorBox(true);
            sheet.addValidationData(validation1);

            CellRangeAddressList addressList2 = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation validation2 = validationHelper.createValidation(constraint, addressList2);
            validation2.setSuppressDropDownArrow(true);
            validation2.setShowErrorBox(true);
            sheet.addValidationData(validation2);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutBuildingDistancesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }

    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "ID_B_DISTANCE",
            "BUILDING_NAME_1",
            "BUILDING_NAME_2",
            "DISTANCE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
        	
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("BUILDING_DISTANCE TEMPLATE");
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

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("BuildingNames");

            CellRangeAddressList addressList1 = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation1 = validationHelper.createValidation(constraint, addressList1);
            validation1.setSuppressDropDownArrow(true);
            validation1.setShowErrorBox(true);
            sheet.addValidationData(validation1);

            CellRangeAddressList addressList2 = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation validation2 = validationHelper.createValidation(constraint, addressList2);
            validation2.setSuppressDropDownArrow(true);
            validation2.setShowErrorBox(true);
            sheet.addValidationData(validation2);

            workbook.write(out);

            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
