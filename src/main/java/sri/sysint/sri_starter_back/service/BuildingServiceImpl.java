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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
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
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.PlantRepo;


@Service
@Transactional
public class BuildingServiceImpl {
	@Autowired
    private BuildingRepo buildingRepo;
	@Autowired
    private PlantRepo plantRepo;
	
    public BuildingServiceImpl(BuildingRepo buildingRepo){
        this.buildingRepo = buildingRepo;
    }
    
    public BigDecimal getNewId() {
    	return buildingRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<Building> getAllBuilding() {
    	Iterable<Building> buildings = buildingRepo.getDataOrderId();
        List<Building> buildingList = new ArrayList<>();
        for (Building item : buildings) {
        	Building buildingTemp = new Building(item);
        	buildingList.add(buildingTemp);
        }
        
        return buildingList;
    }
    
    public Optional<Building> getBuildingById(BigDecimal id) {
    	Optional<Building> building = buildingRepo.findById(id);
    	return building;
    }
    
    public Building saveBuilding(Building building) {
        try {
        	building.setBUILDING_ID(getNewId());
        	building.setSTATUS(BigDecimal.valueOf(1));
        	building.setCREATION_DATE(new Date());
        	building.setLAST_UPDATE_DATE(new Date());
            return buildingRepo.save(building);
        } catch (Exception e) {
            System.err.println("Error saving building: " + e.getMessage());
            throw e;
        }
    }
    
    public Building updateBuilding(Building building) {
        try {
            Optional<Building> currentBuildingOpt = buildingRepo.findById(building.getBUILDING_ID());
            
            if (currentBuildingOpt.isPresent()) {
                Building currentBuilding = currentBuildingOpt.get();
                
                currentBuilding.setBUILDING_NAME(building.getBUILDING_NAME());
                currentBuilding.setPLANT_ID(building.getPLANT_ID());
                currentBuilding.setLAST_UPDATE_DATE(new Date());
                currentBuilding.setLAST_UPDATED_BY(building.getLAST_UPDATED_BY());
                
                return buildingRepo.save(currentBuilding);
            } else {
                throw new RuntimeException("Building with ID " + building.getBUILDING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Buildinga: " + e.getMessage());
            throw e;
        }
    }
    
    public Building deleteBuilding(Building building) {
        try {
            Optional<Building> currentBuildingOpt = buildingRepo.findById(building.getBUILDING_ID());
            
            if (currentBuildingOpt.isPresent()) {
                Building currentBuilding = currentBuildingOpt.get();
                
                currentBuilding.setSTATUS(BigDecimal.valueOf(0));
                currentBuilding.setLAST_UPDATE_DATE(new Date());
                currentBuilding.setLAST_UPDATED_BY(building.getLAST_UPDATED_BY());
                
                return buildingRepo.save(currentBuilding);
            } else {
                throw new RuntimeException("Building with ID " + building.getBUILDING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating building: " + e.getMessage());
            throw e;
        }
    }
    
    public Building restoreBuilding(Building building) {
        try {
            Optional<Building> currentBuildingOpt = buildingRepo.findById(building.getBUILDING_ID());
            
            if (currentBuildingOpt.isPresent()) {
                Building currentBuilding = currentBuildingOpt.get();
                
                currentBuilding.setSTATUS(BigDecimal.valueOf(1)); 
                currentBuilding.setLAST_UPDATE_DATE(new Date());
                currentBuilding.setLAST_UPDATED_BY(building.getLAST_UPDATED_BY());
                
                return buildingRepo.save(currentBuilding);
            } else {
                throw new RuntimeException("Building with ID " + building.getBUILDING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring building: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllBuilding() {
    	buildingRepo.deleteAll();
    }
    
    public boolean isPlantIdExists(BigDecimal plantId) {
        return plantRepo.existsById(plantId);
    }
    
    public ByteArrayInputStream exportBuildingsExcel() throws IOException {
        List<Building> buildings = buildingRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(buildings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<Building> buildings) throws IOException {
        String[] header = {
            "NOMOR",
            "BUILDING_ID",
            "PLANT_NAME",
            "BUILDING_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Plant> activePlants = plantRepo.findPlantActive();
            List<String> plantNames = activePlants.stream()
                .map(Plant::getPLANT_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("BUILDING DATA");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_PLANTS");
            for (int i = 0; i < plantNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(plantNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("PlantNames");
            namedRange.setRefersToFormula("HIDDEN_PLANTS!$A$1:$A$" + plantNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            for (Building b : buildings) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(b.getBUILDING_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell plantNameCell = dataRow.createCell(2);
                plantNameCell.setCellStyle(borderStyle);
                
                String plantName = "";
                if (b.getPLANT_ID() != null) {
                    Plant plant = activePlants.stream()
                        .filter(p -> p.getPLANT_ID().equals(b.getPLANT_ID()))
                        .findFirst()
                        .orElse(null);
                    plantName = (plant != null) ? plant.getPLANT_NAME() : "";
                }
                plantNameCell.setCellValue(plantName);

                Cell nameCell = dataRow.createCell(3);
                nameCell.setCellValue(b.getBUILDING_NAME());
                nameCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("PlantNames");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Building");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
    
    public ByteArrayInputStream layoutBuildingsExcel() throws IOException {
        List<Building> buildings = buildingRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = layoutBuildingExcel();
        return byteArrayInputStream;
    }
    
    public ByteArrayInputStream layoutBuildingExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "BUILDING_ID",
            "PLANT_NAME",
            "BUILDING_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Plant> activePlants = plantRepo.findPlantActive();
            List<String> plantNames = activePlants.stream()
                .map(Plant::getPLANT_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("BUILDING DATA");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_PLANTS");
            for (int i = 0; i < plantNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(plantNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("PlantNames");
            namedRange.setRefersToFormula("HIDDEN_PLANTS!$A$1:$A$" + plantNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("PlantNames");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal membuat layout Building");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}