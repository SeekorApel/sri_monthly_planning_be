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
import sri.sysint.sri_starter_back.model.Quadrant;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.QuadrantRepo;


@Service
@Transactional
public class QuadrantServiceImpl {
	@Autowired
    private QuadrantRepo quadrantRepo;
	
	@Autowired
    private BuildingRepo buildingRepo;
	
    public QuadrantServiceImpl(QuadrantRepo quadrantRepo){
        this.quadrantRepo = quadrantRepo;
    }
    
    public BigDecimal getNewId() {
    	return quadrantRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<Quadrant> getAllQuadrant() {
    	Iterable<Quadrant> quadrants = quadrantRepo.getDataOrderId();
        List<Quadrant> quadrantList = new ArrayList<>();
        for (Quadrant item : quadrants) {
        	Quadrant quadrantTemp = new Quadrant(item);
        	quadrantList.add(quadrantTemp);
        }
        
        return quadrantList;
    }
    
    public Optional<Quadrant> getQuadrantById(BigDecimal id) {
    	Optional<Quadrant> quadrant = quadrantRepo.findById(id);
    	return quadrant;
    }
    
    public Quadrant saveQuadrant(Quadrant quadrant) {
        try {
        	quadrant.setQUADRANT_ID(getNewId());
        	quadrant.setSTATUS(BigDecimal.valueOf(1));
        	quadrant.setCREATION_DATE(new Date());
        	quadrant.setLAST_UPDATE_DATE(new Date());
            return quadrantRepo.save(quadrant);
        } catch (Exception e) {
            System.err.println("Error saving Quadrant: " + e.getMessage());
            throw e;
        }
    }
    
    public Quadrant updateQuadrant(Quadrant quadrant) {
        try {
            Optional<Quadrant> currentQuadrantOpt = quadrantRepo.findById(quadrant.getQUADRANT_ID());
            
            if (currentQuadrantOpt.isPresent()) {
                Quadrant currentQuadrant = currentQuadrantOpt.get();
                
                currentQuadrant.setQUADRANT_NAME(quadrant.getQUADRANT_NAME());
                currentQuadrant.setBUILDING_ID(quadrant.getBUILDING_ID());
                currentQuadrant.setLAST_UPDATE_DATE(new Date());
                currentQuadrant.setLAST_UPDATED_BY(quadrant.getLAST_UPDATED_BY());
                
                return quadrantRepo.save(currentQuadrant);
            } else {
                throw new RuntimeException("Quadrant with ID " + quadrant.getQUADRANT_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Quadrant: " + e.getMessage());
            throw e;
        }
    }
    
    public Quadrant deleteQuadrant(Quadrant quadrant) {
        try {
            Optional<Quadrant> currentQuadrantOpt = quadrantRepo.findById(quadrant.getQUADRANT_ID());
            
            if (currentQuadrantOpt.isPresent()) {
            	Quadrant currentQuadrant = currentQuadrantOpt.get();
                
            	currentQuadrant.setSTATUS(BigDecimal.valueOf(0));
            	currentQuadrant.setLAST_UPDATE_DATE(new Date());
            	currentQuadrant.setLAST_UPDATED_BY(quadrant.getLAST_UPDATED_BY());
                
            	return quadrantRepo.save(currentQuadrant);
            } else {
                throw new RuntimeException("Quadrant with ID " + quadrant.getQUADRANT_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Quadrant: " + e.getMessage());
            throw e;
        }
    }
    
    public Quadrant restoreQuadrant(Quadrant quadrant) {
        try {
            Optional<Quadrant> currentQuadrantOpt = quadrantRepo.findById(quadrant.getQUADRANT_ID());
            
            if (currentQuadrantOpt.isPresent()) {
                Quadrant currentQuadrant = currentQuadrantOpt.get();
                
                currentQuadrant.setSTATUS(BigDecimal.valueOf(1)); 
                currentQuadrant.setLAST_UPDATE_DATE(new Date());
                currentQuadrant.setLAST_UPDATED_BY(quadrant.getLAST_UPDATED_BY());
                
                return quadrantRepo.save(currentQuadrant);
            } else {
                throw new RuntimeException("Quadrant with ID " + quadrant.getQUADRANT_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring Quadrant: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllQuadrant() {
    	quadrantRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportQuadrantsExcel() throws IOException {
        List<Quadrant> quadrants  = quadrantRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(quadrants );
        return byteArrayInputStream;
    }
    
    public ByteArrayInputStream dataToExcel(List<Quadrant> quadrants) throws IOException {
        String[] header = {
            "NOMOR",
            "QUADRANT_ID",
            "BUILDING_NAME",
            "QUADRANT_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("QUADRANT DATA");
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
            for (Quadrant q : quadrants) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(q.getQUADRANT_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                String buildingName = null;
                if (q.getBUILDING_ID() != null) {
                    Building building = buildingRepo.findById(q.getBUILDING_ID()).orElse(null);
                    if (building != null) {
                        buildingName = building.getBUILDING_NAME();
                    }
                }

                Cell buildingNameCell = dataRow.createCell(2);
                buildingNameCell.setCellValue(buildingName != null ? buildingName : "");
                buildingNameCell.setCellStyle(borderStyle);

                Cell nameCell = dataRow.createCell(3);
                nameCell.setCellValue(q.getQUADRANT_NAME());
                nameCell.setCellStyle(borderStyle);
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
            System.out.println("Failed to export Quadrant data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutQuadrantsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel( );
        return byteArrayInputStream;
    }
    
    public ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "QUADRANT_ID",
            "BUILDING_NAME",
            "QUADRANT_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Building> activeBuildings = buildingRepo.findBuildingActive();
            List<String> buildingNames = activeBuildings.stream()
                .map(Building::getBUILDING_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("QUADRANT DATA");
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
            int nomor = 1;
            
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
            System.out.println("Failed to export Quadrant data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}