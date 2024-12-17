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
import sri.sysint.sri_starter_back.model.CuringSize;
import sri.sysint.sri_starter_back.model.Size;

import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.repository.CuringSizeRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;

@Service
@Transactional
public class CuringSizeServiceImpl {
	@Autowired
    private CuringSizeRepo curingSizeRepo;
	@Autowired
    private MachineCuringTypeRepo machineCuringTypeRepo;
	@Autowired
    private SizeRepo sizeRepo;
	
    public CuringSizeServiceImpl(CuringSizeRepo curingSizeRepo){
        this.curingSizeRepo = curingSizeRepo;
    }
    
    public BigDecimal getNewId() {
    	return curingSizeRepo.getNewId().add(BigDecimal.valueOf(1));
    }
    
    public List<CuringSize> getAllCuringSize() {
    	Iterable<CuringSize> curingSizes = curingSizeRepo.getDataOrderId();
        List<CuringSize> curingSizeList = new ArrayList<>();
        for (CuringSize item : curingSizes) {
        	CuringSize curingSizeTemp = new CuringSize(item);
        	curingSizeList.add(curingSizeTemp);
        }
        
        return curingSizeList;
    }
    
    public Optional<CuringSize> getCuringSizeById(BigDecimal id) {
    	Optional<CuringSize> curingSize = curingSizeRepo.findById(id);
    	return curingSize;
    }
    
    
    public CuringSize saveCuringSize(CuringSize curingSize) {
        try {
        	curingSize.setCURINGSIZE_ID(getNewId());
            curingSize.setSTATUS(BigDecimal.valueOf(1));
            curingSize.setCREATION_DATE(new Date());
            curingSize.setLAST_UPDATE_DATE(new Date());
            return curingSizeRepo.save(curingSize);
        } catch (Exception e) {
            System.err.println("Error saving curingSize: " + e.getMessage());
            throw e;
        }
    }
    
    public CuringSize updateCuringSize(CuringSize curingSize) {
        try {
            Optional<CuringSize> currentCuringSizeOpt = curingSizeRepo.findById(curingSize.getCURINGSIZE_ID());
            
            if (currentCuringSizeOpt.isPresent()) {
                CuringSize currentCuringSize = currentCuringSizeOpt.get();
                
                currentCuringSize.setMACHINECURINGTYPE_ID(curingSize.getMACHINECURINGTYPE_ID());
                currentCuringSize.setSIZE_ID(curingSize.getSIZE_ID());
                currentCuringSize.setCAPACITY(curingSize.getCAPACITY());
                currentCuringSize.setLAST_UPDATE_DATE(new Date());
                currentCuringSize.setLAST_UPDATED_BY(curingSize.getLAST_UPDATED_BY());
                
                return curingSizeRepo.save(currentCuringSize);
            } else {
                throw new RuntimeException("Curing Size with ID " + curingSize.getMACHINECURINGTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Curing Size: " + e.getMessage());
            throw e;
        }
    }
    
    public CuringSize deleteCuringSize(CuringSize curingSize) {
        try {
            Optional<CuringSize> currentCuringSizeOpt = curingSizeRepo.findById(curingSize.getCURINGSIZE_ID());
            
            if (currentCuringSizeOpt.isPresent()) {
            	CuringSize currentCuringSize = currentCuringSizeOpt.get();
                
            	currentCuringSize.setSTATUS(BigDecimal.valueOf(0));
            	currentCuringSize.setLAST_UPDATE_DATE(new Date());
            	currentCuringSize.setLAST_UPDATED_BY(curingSize.getLAST_UPDATED_BY());
                
                return curingSizeRepo.save(currentCuringSize);
            } else {
                throw new RuntimeException("Curing Size with ID " + curingSize.getMACHINECURINGTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Curing Size: " + e.getMessage());
            throw e;
        }
    }
    public CuringSize activateCuringSize(CuringSize curingSize) {
        try {
            Optional<CuringSize> currentCuringSizeOpt = curingSizeRepo.findById(curingSize.getCURINGSIZE_ID());
            
            if (currentCuringSizeOpt.isPresent()) {
            	CuringSize currentCuringSize = currentCuringSizeOpt.get();
                
            	currentCuringSize.setSTATUS(BigDecimal.valueOf(1));
            	currentCuringSize.setLAST_UPDATE_DATE(new Date());
            	currentCuringSize.setLAST_UPDATED_BY(curingSize.getLAST_UPDATED_BY());
                
                return curingSizeRepo.save(currentCuringSize);
            } else {
                throw new RuntimeException("Curing Size with ID " + curingSize.getMACHINECURINGTYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Curing Size: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllCuringSize() {
    	curingSizeRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportCuringSizesExcel() throws IOException {
        List<CuringSize> curingSizes = curingSizeRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(curingSizes);
        return byteArrayInputStream;
    }

    private ByteArrayInputStream dataToExcel(List<CuringSize> curingSizes) throws IOException {
        String[] header = {
            "NOMOR",
            "CURING_SIZE_ID",
            "MACHINECURINGTYPE",
            "SIZE",
            "CAPACITY"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<String> machineCuringTypes = machineCuringTypeRepo.findMachineCuringTypeActive().stream()
                .map(MachineCuringType::getMACHINECURINGTYPE_ID)  
                .collect(Collectors.toList());

            List<String> sizes = sizeRepo.findSizeActive().stream()
                .map(Size::getSIZE_ID)  
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("CURING SIZE DATA");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_MACHINE_CURING_TYPES");
            for (int i = 0; i < machineCuringTypes.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineCuringTypes.get(i));
            }

            Name namedRangeMachineCuringTypes = workbook.createName();
            namedRangeMachineCuringTypes.setNameName("MachineCuringTypes");
            namedRangeMachineCuringTypes.setRefersToFormula("HIDDEN_MACHINE_CURING_TYPES!$A$1:$A$" + machineCuringTypes.size());

            Sheet hiddenSheetSizes = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < sizes.size(); i++) {
                Row row = hiddenSheetSizes.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(sizes.get(i));
            }

            Name namedRangeSizes = workbook.createName();
            namedRangeSizes.setNameName("Sizes");
            namedRangeSizes.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + sizes.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetSizes), true);

            int rowIndex = 1;
            for (CuringSize cs : curingSizes) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(cs.getCURINGSIZE_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell machineCuringTypeCell = dataRow.createCell(2);
                machineCuringTypeCell.setCellStyle(borderStyle);

                String machineCuringType = cs.getMACHINECURINGTYPE_ID() != null 
                    ? cs.getMACHINECURINGTYPE_ID().toString() 
                    : "";

                if (machineCuringTypes.contains(machineCuringType)) {
                    machineCuringTypeCell.setCellValue(machineCuringType);
                } else {
                    machineCuringTypeCell.setCellValue("");
                }

                Cell sizeCell = dataRow.createCell(3);
                sizeCell.setCellStyle(borderStyle);
                sizeCell.setCellValue(cs.getSIZE_ID() != null ? cs.getSIZE_ID().toString() : "");

                Cell capacityCell = dataRow.createCell(4);
                capacityCell.setCellValue(cs.getCAPACITY() != null ? cs.getCAPACITY().doubleValue() : 0);
                capacityCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraintMachineCuringTypes = validationHelper.createFormulaListConstraint("MachineCuringTypes");
            CellRangeAddressList addressListMachineCuringTypes = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validationMachineCuringTypes = validationHelper.createValidation(constraintMachineCuringTypes, addressListMachineCuringTypes);
            validationMachineCuringTypes.setSuppressDropDownArrow(true);
            validationMachineCuringTypes.setShowErrorBox(true);
            sheet.addValidationData(validationMachineCuringTypes);

            DataValidationConstraint constraintSizes = validationHelper.createFormulaListConstraint("Sizes");
            CellRangeAddressList addressListSizes = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation validationSizes = validationHelper.createValidation(constraintSizes, addressListSizes);
            validationSizes.setSuppressDropDownArrow(true);
            validationSizes.setShowErrorBox(true);
            sheet.addValidationData(validationSizes);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export Curing Size data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutCuringSizesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }

    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "CURING_SIZE_ID",
            "MACHINECURINGTYPE",
            "SIZE",
            "CAPACITY"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<String> machineCuringTypes = machineCuringTypeRepo.findMachineCuringTypeActive().stream()
                .map(MachineCuringType::getMACHINECURINGTYPE_ID)  
                .collect(Collectors.toList());

            List<String> sizes = sizeRepo.findSizeActive().stream()
                .map(Size::getSIZE_ID)  
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("CURING SIZE DATA");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_MACHINE_CURING_TYPES");
            for (int i = 0; i < machineCuringTypes.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineCuringTypes.get(i));
            }

            Name namedRangeMachineCuringTypes = workbook.createName();
            namedRangeMachineCuringTypes.setNameName("MachineCuringTypes");
            namedRangeMachineCuringTypes.setRefersToFormula("HIDDEN_MACHINE_CURING_TYPES!$A$1:$A$" + machineCuringTypes.size());

            Sheet hiddenSheetSizes = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < sizes.size(); i++) {
                Row row = hiddenSheetSizes.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(sizes.get(i));
            }

            Name namedRangeSizes = workbook.createName();
            namedRangeSizes.setNameName("Sizes");
            namedRangeSizes.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + sizes.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetSizes), true);

            int rowIndex = 1;
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraintMachineCuringTypes = validationHelper.createFormulaListConstraint("MachineCuringTypes");
            CellRangeAddressList addressListMachineCuringTypes = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validationMachineCuringTypes = validationHelper.createValidation(constraintMachineCuringTypes, addressListMachineCuringTypes);
            validationMachineCuringTypes.setSuppressDropDownArrow(true);
            validationMachineCuringTypes.setShowErrorBox(true);
            sheet.addValidationData(validationMachineCuringTypes);

            DataValidationConstraint constraintSizes = validationHelper.createFormulaListConstraint("Sizes");
            CellRangeAddressList addressListSizes = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation validationSizes = validationHelper.createValidation(constraintSizes, addressListSizes);
            validationSizes.setSuppressDropDownArrow(true);
            validationSizes.setShowErrorBox(true);
            sheet.addValidationData(validationSizes);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export Curing Size data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
