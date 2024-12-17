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
import org.apache.poi.ss.usermodel.Name;
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
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.model.TassSize;
import sri.sysint.sri_starter_back.repository.MachineTassTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;
import sri.sysint.sri_starter_back.repository.TassSizeRepo;

@Service
@Transactional
public class TassSizeServiceImpl {
	@Autowired
    private TassSizeRepo tassSizeRepo;
	
	@Autowired
    private SizeRepo sizeRepo;
	
	@Autowired
    private MachineTassTypeRepo machineTassTypeRepo;
	
    public TassSizeServiceImpl(TassSizeRepo tassSizeRepo){
        this.tassSizeRepo = tassSizeRepo;
    }
    
    public BigDecimal getNewId() {
    	return tassSizeRepo.getNewId().add(BigDecimal.valueOf(1));
    }
    
    public List<TassSize> getAllTassSize() {
    	Iterable<TassSize> tassSize = tassSizeRepo.getDataOrderId();
        List<TassSize> tassSizeList = new ArrayList<>();
        for (TassSize item : tassSize) {
        	TassSize curingSizeTemp = new TassSize(item);
        	tassSizeList.add(curingSizeTemp);
        }
        
        return tassSizeList;
    }
    
    public Optional<TassSize> getTassSizeById(BigDecimal id) {
    	Optional<TassSize> tassSize = tassSizeRepo.findById(id);
    	return tassSize;
    }
    
    
    public TassSize saveTassSize(TassSize tassSize) {
        try {
        	tassSize.setTASSIZE_ID(getNewId());
        	tassSize.setSTATUS(BigDecimal.valueOf(1));
        	tassSize.setCREATION_DATE(new Date());
        	tassSize.setLAST_UPDATE_DATE(new Date());
            return tassSizeRepo.save(tassSize);
        } catch (Exception e) {
            System.err.println("Error saving Tass Size: " + e.getMessage());
            throw e;
        }
    }
    
    public TassSize updateTassSize(TassSize tassSize) {
        try {
            Optional<TassSize> currentTassSizeOpt = tassSizeRepo.findById(tassSize.getTASSIZE_ID());
            
            if (currentTassSizeOpt.isPresent()) {
            	TassSize currentTassSize = currentTassSizeOpt.get();
                
            	currentTassSize.setSIZE_ID(tassSize.getSIZE_ID());
                currentTassSize.setCAPACITY(tassSize.getCAPACITY());
                currentTassSize.setMACHINETASSTYPE_ID(tassSize.getMACHINETASSTYPE_ID());
                
                currentTassSize.setLAST_UPDATE_DATE(new Date());
                currentTassSize.setLAST_UPDATED_BY(tassSize.getLAST_UPDATED_BY());
                
                return tassSizeRepo.save(currentTassSize);
            } else {
                throw new RuntimeException("Tass Size with ID " + tassSize.getTASSIZE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Tass Size: " + e.getMessage());
            throw e;
        }
    }
    
    public TassSize deleteTassSize(TassSize tassSize) {
        try {
            Optional<TassSize> currentTassSizeOpt = tassSizeRepo.findById(tassSize.getTASSIZE_ID());
            
            if (currentTassSizeOpt.isPresent()) {
            	TassSize currentTassSize = currentTassSizeOpt.get();
                
            	currentTassSize.setSTATUS(BigDecimal.valueOf(0));
            	currentTassSize.setLAST_UPDATE_DATE(new Date());
            	currentTassSize.setLAST_UPDATED_BY(tassSize.getLAST_UPDATED_BY());
                
                return tassSizeRepo.save(currentTassSize);
            } else {
                throw new RuntimeException("Tass Size with ID " + tassSize.getMACHINETASSTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Tass Size: " + e.getMessage());
            throw e;
        }
    }
    
    public TassSize activateTassSize(TassSize tassSize) {
        try {
            Optional<TassSize> currentTassSizeOpt = tassSizeRepo.findById(tassSize.getTASSIZE_ID());
            
            if (currentTassSizeOpt.isPresent()) {
            	TassSize currentTassSize = currentTassSizeOpt.get();
                
            	currentTassSize.setSTATUS(BigDecimal.valueOf(1));
            	currentTassSize.setLAST_UPDATE_DATE(new Date());
            	currentTassSize.setLAST_UPDATED_BY(tassSize.getLAST_UPDATED_BY());
                
                return tassSizeRepo.save(currentTassSize);
            } else {
                throw new RuntimeException("Tass Size with ID " + tassSize.getMACHINETASSTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Tass Size: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllTassSize() {
    	tassSizeRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportTassSizeExcel() throws IOException {
        List<TassSize> tassSizes = tassSizeRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(tassSizes);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<TassSize> tassSizes) throws IOException {
        String[] header = { 
            "NOMOR", 
            "TASSIZE_ID", 
            "MACHINETASSTYPE", 
            "SIZE", 
            "CAPACITY" 
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<Size> activeSizes = sizeRepo.findSizeActive();
            List<MachineTassType> activeTassTypes = machineTassTypeRepo.findMachineTassTypeActive();
            List<String> sizeIDs = activeSizes.stream()
                .map(Size::getSIZE_ID)
                .collect(Collectors.toList());
             List<String> tassTypeIDs = activeTassTypes.stream()
                 .map(MachineTassType::getMACHINETASSTYPE_ID)
                 .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("TASSIZE DATA");
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

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 1;
            for (TassSize ts : tassSizes) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell tassSizeIdCell = dataRow.createCell(1);
                tassSizeIdCell.setCellValue(ts.getTASSIZE_ID().doubleValue());
                tassSizeIdCell.setCellStyle(borderStyle);

                Cell machineTassTypeIdCell = dataRow.createCell(2);
                machineTassTypeIdCell.setCellValue(ts.getMACHINETASSTYPE_ID() != null ? ts.getMACHINETASSTYPE_ID() : null);
                machineTassTypeIdCell.setCellStyle(borderStyle);

                Cell sizeIdCell = dataRow.createCell(3);
                sizeIdCell.setCellValue(ts.getSIZE_ID() != null ? ts.getSIZE_ID() : null);
                sizeIdCell.setCellStyle(borderStyle);

                Cell capacityCell = dataRow.createCell(4);
                capacityCell.setCellValue(ts.getCAPACITY() != null ? ts.getCAPACITY().doubleValue() : null);
                capacityCell.setCellStyle(borderStyle);
            }
            
            Sheet hiddenSheetTassType = workbook.createSheet("HIDDEN_TASSTYPES");
            for (int i = 0; i < tassTypeIDs.size(); i++) {
                Row row = hiddenSheetTassType.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(tassTypeIDs.get(i));
            }

            Name namedRangeTassType = workbook.createName();
            namedRangeTassType.setNameName("tassTypeIDs");
            namedRangeTassType.setRefersToFormula("HIDDEN_TASSTYPES!$A$1:$A$" + tassTypeIDs.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetTassType), true);

            Sheet hiddenSheetSize = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < sizeIDs.size(); i++) {
                Row row = hiddenSheetSize.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(sizeIDs.get(i));
            }

            Name namedRangeSize = workbook.createName();
            namedRangeSize.setNameName("sizeIDs");
            namedRangeSize.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + sizeIDs.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetSize), true);

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint tassTypeConstraint = validationHelper.createFormulaListConstraint("tassTypeIDs");
            CellRangeAddressList tassTypeAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation tassTypeValidation = validationHelper.createValidation(tassTypeConstraint, tassTypeAddressList);
            tassTypeValidation.setSuppressDropDownArrow(true);
            tassTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(tassTypeValidation);

            DataValidationConstraint sizeConstraint = validationHelper.createFormulaListConstraint("sizeIDs");
            CellRangeAddressList sizeAddressList = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation sizeValidation = validationHelper.createValidation(sizeConstraint, sizeAddressList);
            sizeValidation.setSuppressDropDownArrow(true);
            sizeValidation.setShowErrorBox(true);
            sheet.addValidationData(sizeValidation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export Tass Size data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutTassSizesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }

    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = { 
            "NOMOR", 
            "TASSIZE_ID", 
            "MACHINETASSTYPE", 
            "SIZE", 
            "CAPACITY" 
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            List<Size> activeSizes = sizeRepo.findSizeActive();
            List<MachineTassType> activeTassTypes = machineTassTypeRepo.findMachineTassTypeActive();
            List<String> sizeIDs = activeSizes.stream()
                .map(Size::getSIZE_ID)
                .collect(Collectors.toList());
             List<String> tassTypeIDs = activeTassTypes.stream()
                 .map(MachineTassType::getMACHINETASSTYPE_ID)
                 .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("TASS SIZE DATA");
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

            Sheet hiddenSheetTassType = workbook.createSheet("HIDDEN_TASSTYPES");
            for (int i = 0; i < tassTypeIDs.size(); i++) {
                Row row = hiddenSheetTassType.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(tassTypeIDs.get(i));
            }

            Name namedRangeTassType = workbook.createName();
            namedRangeTassType.setNameName("tassTypeIDs");
            namedRangeTassType.setRefersToFormula("HIDDEN_TASSTYPES!$A$1:$A$" + tassTypeIDs.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetTassType), true);

            Sheet hiddenSheetSize = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < sizeIDs.size(); i++) {
                Row row = hiddenSheetSize.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(sizeIDs.get(i));
            }

            Name namedRangeSize = workbook.createName();
            namedRangeSize.setNameName("sizeIDs");
            namedRangeSize.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + sizeIDs.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetSize), true);

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint tassTypeConstraint = validationHelper.createFormulaListConstraint("tassTypeIDs");
            CellRangeAddressList tassTypeAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation tassTypeValidation = validationHelper.createValidation(tassTypeConstraint, tassTypeAddressList);
            tassTypeValidation.setSuppressDropDownArrow(true);
            tassTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(tassTypeValidation);

            DataValidationConstraint sizeConstraint = validationHelper.createFormulaListConstraint("sizeIDs");
            CellRangeAddressList sizeAddressList = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation sizeValidation = validationHelper.createValidation(sizeConstraint, sizeAddressList);
            sizeValidation.setSuppressDropDownArrow(true);
            sizeValidation.setShowErrorBox(true);
            sheet.addValidationData(sizeValidation);

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
