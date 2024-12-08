package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
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
import sri.sysint.sri_starter_back.model.MaxCapacity;
import sri.sysint.sri_starter_back.repository.MaxCapacityRepo;


@Service
@Transactional
public class MaxCapacityServiceImpl {
	@Autowired
	private MaxCapacityRepo maxCapacityRepo;
	
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
            "MACHINECURINGTYPE_ID",
            "CYCLE_TIME",
            "CAPACITY_SHIFT_1",
            "CAPACITY_SHIFT_2",
            "CAPACITY_SHIFT_3"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
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

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256); // Set column width
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 1;
            for (MaxCapacity m : maxCapacities) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1); // Set the NOMOR cell
                nomorCell.setCellStyle(borderStyle);

                Cell maxCapIdCell = dataRow.createCell(1);
                maxCapIdCell.setCellValue(m.getMAX_CAP_ID() != null ? m.getMAX_CAP_ID().doubleValue() : null);
                maxCapIdCell.setCellStyle(borderStyle);

                Cell productIdCell = dataRow.createCell(2);
                productIdCell.setCellValue(m.getPRODUCT_ID() != null ? m.getPRODUCT_ID().doubleValue() : null);
                productIdCell.setCellStyle(borderStyle);

                Cell machineCuringTypeIdCell = dataRow.createCell(3);
                machineCuringTypeIdCell.setCellValue(m.getMACHINECURINGTYPE_ID() != null ? m.getMACHINECURINGTYPE_ID() : "");
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