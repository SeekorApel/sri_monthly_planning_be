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
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Quadrant;
import sri.sysint.sri_starter_back.model.QuadrantDistance;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.QuadrantDistanceRepo;
import sri.sysint.sri_starter_back.repository.QuadrantRepo;

@Service
@Transactional
public class QuadrantDistanceServiceImpl {
    @Autowired
    private QuadrantDistanceRepo quadrantDistanceRepo;
    
    @Autowired
    private QuadrantRepo quadrantRepo;
    
    public QuadrantDistanceServiceImpl(QuadrantDistanceRepo quadrantDistanceRepo){
        this.quadrantDistanceRepo = quadrantDistanceRepo;
    }
    
    public BigDecimal getNewId() {
        return quadrantDistanceRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<QuadrantDistance> getAllQuadrantDistance() {
        Iterable<QuadrantDistance> quadrantDistances = quadrantDistanceRepo.getDataOrderId();
        List<QuadrantDistance> quadrantDistanceList = new ArrayList<>();
        for (QuadrantDistance item : quadrantDistances) {
            QuadrantDistance quadrantDistanceTemp = new QuadrantDistance(item);
            quadrantDistanceList.add(quadrantDistanceTemp);
        }
        
        return quadrantDistanceList;
    }
    
    public Optional<QuadrantDistance> getQuadrantDistanceById(BigDecimal id) {
        Optional<QuadrantDistance> quadrantDistance = quadrantDistanceRepo.findById(id);
        return quadrantDistance;
    }
    
    public QuadrantDistance saveQuadrantDistance(QuadrantDistance quadrantDistance) {
        try {
            quadrantDistance.setID_Q_DISTANCE(getNewId());
            quadrantDistance.setSTATUS(BigDecimal.valueOf(1));
            quadrantDistance.setCREATION_DATE(new Date());
            quadrantDistance.setLAST_UPDATE_DATE(new Date());
            return quadrantDistanceRepo.save(quadrantDistance);
        } catch (Exception e) {
            System.err.println("Error saving quadrantDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public QuadrantDistance updateQuadrantDistance(QuadrantDistance quadrantDistance) {
        try {
            Optional<QuadrantDistance> currentQuadrantDistanceOpt = quadrantDistanceRepo.findById(quadrantDistance.getID_Q_DISTANCE());
            
            if (currentQuadrantDistanceOpt.isPresent()) {
                QuadrantDistance currentQuadrantDistance = currentQuadrantDistanceOpt.get();
                
                currentQuadrantDistance.setQUADRANT_ID_1(quadrantDistance.getQUADRANT_ID_1());
                currentQuadrantDistance.setQUADRANT_ID_2(quadrantDistance.getQUADRANT_ID_2());
                currentQuadrantDistance.setDISTANCE(quadrantDistance.getDISTANCE());
                currentQuadrantDistance.setLAST_UPDATE_DATE(new Date());
                currentQuadrantDistance.setLAST_UPDATED_BY(quadrantDistance.getLAST_UPDATED_BY());
                
                return quadrantDistanceRepo.save(currentQuadrantDistance);
            } else {
                throw new RuntimeException("QuadrantDistance with ID " + quadrantDistance.getID_Q_DISTANCE() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating quadrantDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public QuadrantDistance deleteQuadrantDistance(QuadrantDistance quadrantDistance) {
        try {
            Optional<QuadrantDistance> currentQuadrantDistanceOpt = quadrantDistanceRepo.findById(quadrantDistance.getID_Q_DISTANCE());
            
            if (currentQuadrantDistanceOpt.isPresent()) {
                QuadrantDistance currentQuadrantDistance = currentQuadrantDistanceOpt.get();
                
                currentQuadrantDistance.setSTATUS(BigDecimal.valueOf(0));
                currentQuadrantDistance.setLAST_UPDATE_DATE(new Date());
                currentQuadrantDistance.setLAST_UPDATED_BY(quadrantDistance.getLAST_UPDATED_BY());
                
                return quadrantDistanceRepo.save(currentQuadrantDistance);
            } else {
                throw new RuntimeException("QuadrantDistance with ID " + quadrantDistance.getID_Q_DISTANCE() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating quadrantDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public QuadrantDistance restoreQuadrantDistance(QuadrantDistance quadrantDistance) {
        try {
            Optional<QuadrantDistance> currentQuadrantDistanceOpt = quadrantDistanceRepo.findById(quadrantDistance.getID_Q_DISTANCE());
            
            if (currentQuadrantDistanceOpt.isPresent()) {
                QuadrantDistance currentQuadrantDistance = currentQuadrantDistanceOpt.get();
                
                currentQuadrantDistance.setSTATUS(BigDecimal.valueOf(1)); 
                currentQuadrantDistance.setLAST_UPDATE_DATE(new Date());
                currentQuadrantDistance.setLAST_UPDATED_BY(quadrantDistance.getLAST_UPDATED_BY());
                
                return quadrantDistanceRepo.save(currentQuadrantDistance);
            } else {
                throw new RuntimeException("QuadrantDistance with ID " + quadrantDistance.getID_Q_DISTANCE() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring quadrantDistance: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllQuadrantDistance() {
        quadrantDistanceRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportQuadrantDistancesExcel() throws IOException {
        List<QuadrantDistance> quadrantDistances  = quadrantDistanceRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(quadrantDistances );
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<QuadrantDistance> quadrantDistances) throws IOException {
        String[] header = {
            "NOMOR",
            "ID_Q_DISTANCE",
            "QUADRANT_NAME_1", 
            "QUADRANT_NAME_2",  
            "DISTANCE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Quadrant> activeQuadrants = quadrantRepo.findQuadrantActive();
            List<String> quadrantNames = activeQuadrants.stream()
                .map(Quadrant::getQUADRANT_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("QUADRANT_DISTANCE DATA");

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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_QUADRANTS");
            for (int i = 0; i < quadrantNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(quadrantNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("QuadrantNames");
            namedRange.setRefersToFormula("HIDDEN_QUADRANTS!$A$1:$A$" + quadrantNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            for (QuadrantDistance qd : quadrantDistances) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(qd.getID_Q_DISTANCE().doubleValue());
                idCell.setCellStyle(borderStyle);

                String quadrantName1 = null;
                if (qd.getQUADRANT_ID_1() != null) {
                    Quadrant quadrant1 = quadrantRepo.findById(qd.getQUADRANT_ID_1()).orElse(null);
                    if (quadrant1 != null) {
                        quadrantName1 = quadrant1.getQUADRANT_NAME();
                    }
                }

                String quadrantName2 = null;
                if (qd.getQUADRANT_ID_2() != null) {
                    Quadrant quadrant2 = quadrantRepo.findById(qd.getQUADRANT_ID_2()).orElse(null);
                    if (quadrant2 != null) {
                        quadrantName2 = quadrant2.getQUADRANT_NAME();
                    }
                }

                Cell quadrantName1Cell = dataRow.createCell(2);
                quadrantName1Cell.setCellValue(quadrantName1 != null ? quadrantName1 : "Unknown");
                quadrantName1Cell.setCellStyle(borderStyle);

                Cell quadrantName2Cell = dataRow.createCell(3);
                quadrantName2Cell.setCellValue(quadrantName2 != null ? quadrantName2 : "Unknown");
                quadrantName2Cell.setCellStyle(borderStyle);

                Cell distanceCell = dataRow.createCell(4);
                distanceCell.setCellValue(qd.getDISTANCE() != null ? qd.getDISTANCE().doubleValue() : null);
                distanceCell.setCellStyle(borderStyle);
            }
            
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("QuadrantNames");

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

    public ByteArrayInputStream layoutQuadrantDistancesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel( );
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "ID_Q_DISTANCE",
            "QUADRANT_NAME_1", 
            "QUADRANT_NAME_2",  
            "DISTANCE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Quadrant> activeQuadrants = quadrantRepo.findQuadrantActive();
            List<String> quadrantNames = activeQuadrants.stream()
                .map(Quadrant::getQUADRANT_NAME)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("QUADRANT_DISTANCE DATA");

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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_QUADRANTS");
            for (int i = 0; i < quadrantNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(quadrantNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("QuadrantNames");
            namedRange.setRefersToFormula("HIDDEN_QUADRANTS!$A$1:$A$" + quadrantNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("QuadrantNames");

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
