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
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.MaxCapacity;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.repository.MaxCapacityRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;




@Service
@Transactional
public class MaxCapacityServiceImpl {
	@Autowired
	private MaxCapacityRepo maxCapacityRepo;
	@Autowired
	private ProductRepo productRepo;
	@Autowired
	private MachineCuringTypeRepo machineCuringTypeRepo;
	public MaxCapacityServiceImpl(MaxCapacityRepo maxCapacityRepo){
        this.maxCapacityRepo = maxCapacityRepo;
    }
    
    public BigDecimal getNewId() {
        return maxCapacityRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<MaxCapacity> getAllMaxCapacity() {
        Iterable<MaxCapacity> maxCapacities = maxCapacityRepo.getDataOrderId();
        List<MaxCapacity> maxCapacityList = new ArrayList<>();
        for (MaxCapacity item : maxCapacities) {
            MaxCapacity maxCapacityTemp = new MaxCapacity(item);
            maxCapacityList.add(maxCapacityTemp);
        }
        
        return maxCapacityList;
    }
    
    public Optional<MaxCapacity> getMaxCapacityById(BigDecimal id) {
        Optional<MaxCapacity> maxCapacity = maxCapacityRepo.findById(id);
        return maxCapacity;
    }
    
    public MaxCapacity saveMaxCapacity(MaxCapacity maxCapacity) {
        try {
            maxCapacity.setMAX_CAP_ID(getNewId());
            maxCapacity.setSTATUS(BigDecimal.valueOf(1));
            maxCapacity.setCREATION_DATE(new Date());
            maxCapacity.setLAST_UPDATE_DATE(new Date());
            return maxCapacityRepo.save(maxCapacity);
        } catch (Exception e) {
            System.err.println("Error saving maxCapacity: " + e.getMessage());
            throw e;
        }
    }
    
    public MaxCapacity updateMaxCapacity(MaxCapacity maxCapacity) {
        try {
            Optional<MaxCapacity> currentMaxCapacityOpt = maxCapacityRepo.findById(maxCapacity.getMAX_CAP_ID());
            
            if (currentMaxCapacityOpt.isPresent()) {
                MaxCapacity currentMaxCapacity = currentMaxCapacityOpt.get();
                
                currentMaxCapacity.setPRODUCT_ID(maxCapacity.getPRODUCT_ID());
                currentMaxCapacity.setMACHINECURINGTYPE_ID(maxCapacity.getMACHINECURINGTYPE_ID());
                currentMaxCapacity.setCYCLE_TIME(maxCapacity.getCYCLE_TIME());
                currentMaxCapacity.setCAPACITY_SHIFT_1(maxCapacity.getCAPACITY_SHIFT_1());
                currentMaxCapacity.setCAPACITY_SHIFT_2(maxCapacity.getCAPACITY_SHIFT_2());
                currentMaxCapacity.setCAPACITY_SHIFT_3(maxCapacity.getCAPACITY_SHIFT_3());
                currentMaxCapacity.setLAST_UPDATE_DATE(new Date());
                currentMaxCapacity.setLAST_UPDATED_BY(maxCapacity.getLAST_UPDATED_BY());
                
                return maxCapacityRepo.save(currentMaxCapacity);
            } else {
                throw new RuntimeException("MaxCapacity with ID " + maxCapacity.getMAX_CAP_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating maxCapacity: " + e.getMessage());
            throw e;
        }
    }
    
    public MaxCapacity deleteMaxCapacity(MaxCapacity maxCapacity) {
        try {
            Optional<MaxCapacity> currentMaxCapacityOpt = maxCapacityRepo.findById(maxCapacity.getMAX_CAP_ID());
            
            if (currentMaxCapacityOpt.isPresent()) {
                MaxCapacity currentMaxCapacity = currentMaxCapacityOpt.get();
                
                currentMaxCapacity.setSTATUS(BigDecimal.valueOf(0));
                currentMaxCapacity.setLAST_UPDATE_DATE(new Date());
                currentMaxCapacity.setLAST_UPDATED_BY(maxCapacity.getLAST_UPDATED_BY());
                
                return maxCapacityRepo.save(currentMaxCapacity);
            } else {
                throw new RuntimeException("MaxCapacity with ID " + maxCapacity.getMAX_CAP_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating maxCapacity: " + e.getMessage());
            throw e;
        }
    }
    
    public MaxCapacity restoreMaxCapacity(MaxCapacity maxCapacity) {
        try {
            Optional<MaxCapacity> currentMaxCapacityOpt = maxCapacityRepo.findById(maxCapacity.getMAX_CAP_ID());
            
            if (currentMaxCapacityOpt.isPresent()) {
                MaxCapacity currentMaxCapacity = currentMaxCapacityOpt.get();
                
                currentMaxCapacity.setSTATUS(BigDecimal.valueOf(1)); 
                currentMaxCapacity.setLAST_UPDATE_DATE(new Date());
                currentMaxCapacity.setLAST_UPDATED_BY(maxCapacity.getLAST_UPDATED_BY());
                
                return maxCapacityRepo.save(currentMaxCapacity);
            } else {
                throw new RuntimeException("MaxCapacity with ID " + maxCapacity.getMAX_CAP_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring maxCapacity: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllMaxCapacity() {
        maxCapacityRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMaxCapacitysExcel() throws IOException {
        List<MaxCapacity> maxCapacities  = maxCapacityRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(maxCapacities );
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MaxCapacity> maxCapacities) throws IOException {
        String[] header = {
            "NOMOR",
            "MAX_CAPACITY_ID",
            "PART_NUMBER",
            "MACHINECURINGTYPE",
            "CYCLE_TIME",
            "CAPACITY_SHIFT_1",
            "CAPACITY_SHIFT_2",
            "CAPACITY_SHIFT_3"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<BigDecimal> productNames = productRepo.findProductActive()
                .stream().map(Product::getPART_NUMBER).collect(Collectors.toList());
            List<String> machineCuringTypes = machineCuringTypeRepo.findMachineCuringTypeActive()
                .stream().map(MachineCuringType::getMACHINECURINGTYPE_ID).collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MAX CAPACITY DATA");
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

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            Sheet hiddenSheet1 = workbook.createSheet("HIDDEN_PRODUCTS");
            for (int i = 0; i < productNames.size(); i++) {
                Row row = hiddenSheet1.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(productNames.get(i).doubleValue());
            }
            
            Sheet hiddenSheet2 = workbook.createSheet("HIDDEN_MACHINES");
            for (int i = 0; i < machineCuringTypes.size(); i++) {
                Row row = hiddenSheet2.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineCuringTypes.get(i));
            }

            Name namedRange1 = workbook.createName();
            namedRange1.setNameName("ProductNames");
            namedRange1.setRefersToFormula("HIDDEN_PRODUCTS!$A$1:$A$" + productNames.size());

            Name namedRange2 = workbook.createName();
            namedRange2.setNameName("MachineTypes");
            namedRange2.setRefersToFormula("HIDDEN_MACHINES!$A$1:$A$" + machineCuringTypes.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet1), true);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet2), true);

            int rowIndex = 1;
            for (MaxCapacity m : maxCapacities) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell maxCapIdCell = dataRow.createCell(1);
                maxCapIdCell.setCellValue(m.getMAX_CAP_ID() != null ? m.getMAX_CAP_ID().doubleValue() : null);
                maxCapIdCell.setCellStyle(borderStyle);
                
                BigDecimal partnumber = null;
                if (m.getPRODUCT_ID() != null) {
                    Product product = productRepo.findById(m.getPRODUCT_ID()).orElse(null);
                    if (product != null) {
                    	partnumber = product.getPART_NUMBER();
                    }
                }
                
                String machineCuringTypeid = null;
                if (m.getMACHINECURINGTYPE_ID() != null) {
                    MachineCuringType machineCuringType = machineCuringTypeRepo.findById(m.getMACHINECURINGTYPE_ID()).orElse(null);
                    if (machineCuringType != null) {
                    	machineCuringTypeid = machineCuringType.getMACHINECURINGTYPE_ID();
                    }
                }

                Cell productIdCell = dataRow.createCell(2);
                productIdCell.setCellValue(partnumber != null ? partnumber.doubleValue() : null);
                productIdCell.setCellStyle(borderStyle);

                Cell machineCuringTypeIdCell = dataRow.createCell(3);
                machineCuringTypeIdCell.setCellValue(machineCuringTypeid != null ? machineCuringTypeid : null);
                machineCuringTypeIdCell.setCellStyle(borderStyle);

                Cell cycleTimeCell = dataRow.createCell(4);
                cycleTimeCell.setCellValue(m.getCYCLE_TIME() != null ? m.getCYCLE_TIME().doubleValue() : null);
                cycleTimeCell.setCellStyle(borderStyle);

                Cell capacityShift1Cell = dataRow.createCell(5);
                capacityShift1Cell.setCellValue(m.getCAPACITY_SHIFT_1() != null ? m.getCAPACITY_SHIFT_1().doubleValue() : null);
                capacityShift1Cell.setCellStyle(borderStyle);

                Cell capacityShift2Cell = dataRow.createCell(6);
                capacityShift2Cell.setCellValue(m.getCAPACITY_SHIFT_2() != null ? m.getCAPACITY_SHIFT_2().doubleValue() : null);
                capacityShift2Cell.setCellStyle(borderStyle);

                Cell capacityShift3Cell = dataRow.createCell(7);
                capacityShift3Cell.setCellValue(m.getCAPACITY_SHIFT_3() != null ? m.getCAPACITY_SHIFT_3().doubleValue() : null);
                capacityShift3Cell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint productConstraint = validationHelper.createFormulaListConstraint("ProductNames");
            CellRangeAddressList productAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation productValidation = validationHelper.createValidation(productConstraint, productAddressList);
            productValidation.setSuppressDropDownArrow(true);
            productValidation.setShowErrorBox(true);
            sheet.addValidationData(productValidation);

            DataValidationConstraint machineConstraint = validationHelper.createFormulaListConstraint("MachineTypes");
            CellRangeAddressList machineAddressList = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation machineValidation = validationHelper.createValidation(machineConstraint, machineAddressList);
            machineValidation.setSuppressDropDownArrow(true);
            machineValidation.setShowErrorBox(true);
            sheet.addValidationData(machineValidation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Max Capacity");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutMaxCapacitysExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "MAX_CAPACITY_ID",
            "PART_NUMBER",
            "MACHINECURINGTYPE",
            "CYCLE_TIME",
            "CAPACITY_SHIFT_1",
            "CAPACITY_SHIFT_2",
            "CAPACITY_SHIFT_3"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<BigDecimal> productNames = productRepo.findProductActive()
                .stream().map(Product::getPART_NUMBER).collect(Collectors.toList());
            List<String> machineCuringTypes = machineCuringTypeRepo.findMachineCuringTypeActive()
                .stream().map(MachineCuringType::getMACHINECURINGTYPE_ID).collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MAX CAPACITY DATA");
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

            Sheet hiddenSheet1 = workbook.createSheet("HIDDEN_PRODUCTS");
            for (int i = 0; i < productNames.size(); i++) {
                Row row = hiddenSheet1.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(productNames.get(i).doubleValue());
            }
            
            Sheet hiddenSheet2 = workbook.createSheet("HIDDEN_MACHINES");
            for (int i = 0; i < machineCuringTypes.size(); i++) {
                Row row = hiddenSheet2.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineCuringTypes.get(i));
            }

            Name namedRange1 = workbook.createName();
            namedRange1.setNameName("ProductNames");
            namedRange1.setRefersToFormula("HIDDEN_PRODUCTS!$A$1:$A$" + productNames.size());

            Name namedRange2 = workbook.createName();
            namedRange2.setNameName("MachineTypes");
            namedRange2.setRefersToFormula("HIDDEN_MACHINES!$A$1:$A$" + machineCuringTypes.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet1), true);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet2), true);

            int rowIndex = 1;
            
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint productConstraint = validationHelper.createFormulaListConstraint("ProductNames");
            CellRangeAddressList productAddressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation productValidation = validationHelper.createValidation(productConstraint, productAddressList);
            productValidation.setSuppressDropDownArrow(true);
            productValidation.setShowErrorBox(true);
            sheet.addValidationData(productValidation);

            DataValidationConstraint machineConstraint = validationHelper.createFormulaListConstraint("MachineTypes");
            CellRangeAddressList machineAddressList = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation machineValidation = validationHelper.createValidation(machineConstraint, machineAddressList);
            machineValidation.setSuppressDropDownArrow(true);
            machineValidation.setShowErrorBox(true);
            sheet.addValidationData(machineValidation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Max Capacity");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }



}