package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
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

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Set column widths
            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256); // 20 characters wide
            }

            // Write header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 1;
            int nomor = 1;  
            for (BuildingDistance bd : buildingDistances) {
                Row dataRow = sheet.createRow(rowIndex++);

                // Write Nomor
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                // Write ID_B_DISTANCE
                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(bd.getID_B_DISTANCE().doubleValue());
                idCell.setCellStyle(borderStyle);

                String buildingName1 = "";
                if (bd.getBUILDING_ID_1() != null) {
                    Optional<Building> building1Opt = buildingRepo.findById(bd.getBUILDING_ID_1());
                    if (building1Opt.isPresent()) {
                        buildingName1 = building1Opt.get().getBUILDING_NAME();
                    }
                }

                Cell buildingName1Cell = dataRow.createCell(2);
                buildingName1Cell.setCellValue(buildingName1);
                buildingName1Cell.setCellStyle(borderStyle);

                String buildingName2 = "";
                if (bd.getBUILDING_ID_2() != null) {
                    Optional<Building> building2Opt = buildingRepo.findById(bd.getBUILDING_ID_2());
                    if (building2Opt.isPresent()) {
                        buildingName2 = building2Opt.get().getBUILDING_NAME();
                    }
                }

                Cell buildingName2Cell = dataRow.createCell(3);
                buildingName2Cell.setCellValue(buildingName2);
                buildingName2Cell.setCellStyle(borderStyle);

                Cell distanceCell = dataRow.createCell(4);
                distanceCell.setCellValue(bd.getDISTANCE() != null ? bd.getDISTANCE().doubleValue() : null);
                distanceCell.setCellStyle(borderStyle);
            }

            // Write the workbook to the output stream
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export BuildingDistance data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }


}
