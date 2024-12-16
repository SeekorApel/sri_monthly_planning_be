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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.CTCuring;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.repository.CTCuringRepo;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;

@Service
@Transactional
public class CTCuringServiceImpl {
	@Autowired
    private CTCuringRepo ctCuringRepo;

    @Autowired
    private ItemCuringRepo itemCuringRepo;

    @Autowired
    private MachineCuringRepo machineCuringRepo;
	
    public CTCuringServiceImpl(CTCuringRepo ctCuringRepo){
        this.ctCuringRepo = ctCuringRepo;
    }
    
    public BigDecimal getNewId() {
    	return ctCuringRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<CTCuring> getAllCTCuring() {
    	Iterable<CTCuring> ctCurings = ctCuringRepo.getDataOrderId();
        List<CTCuring> ctCuringList = new ArrayList<>();
        for (CTCuring item : ctCurings) {
            CTCuring ctCuringTemp = new CTCuring(item);
            ctCuringList.add(ctCuringTemp);
        }
        
        return ctCuringList;
    }
    
    public Optional<CTCuring> getCTCuringById(BigDecimal id) {
    	Optional<CTCuring> ctCuring = ctCuringRepo.findById(id);
    	return ctCuring;
    }
    
    
    public CTCuring saveCTCuring(CTCuring ctCuring) {
        try {
        	ctCuring.setCT_CURING_ID(getNewId());
        	ctCuring.setSTATUS(BigDecimal.valueOf(1));
        	ctCuring.setCREATION_DATE(new Date());
        	ctCuring.setLAST_UPDATE_DATE(new Date());
            return ctCuringRepo.save(ctCuring);
        } catch (Exception e) {
            System.err.println("Error saving CT CURING: " + e.getMessage());
            throw e;
        }
    }
    
    public CTCuring updateCTCuring(CTCuring ctCuring) {
        try {
            Optional<CTCuring> currentCTCuringOpt = ctCuringRepo.findById(ctCuring.getCT_CURING_ID());
            
            if (currentCTCuringOpt.isPresent()) {
                CTCuring currentCTCuring = currentCTCuringOpt.get();
                
             
                currentCTCuring.setWIP(ctCuring.getWIP());
                currentCTCuring.setGROUP_COUNTER(ctCuring.getGROUP_COUNTER());
                currentCTCuring.setVAR_GROUP_COUNTER(ctCuring.getVAR_GROUP_COUNTER());
                currentCTCuring.setSEQUENCE(ctCuring.getSEQUENCE());
                currentCTCuring.setWCT(ctCuring.getWCT());
                currentCTCuring.setOPERATION_SHORT_TEXT(ctCuring.getOPERATION_SHORT_TEXT());
                currentCTCuring.setOPERATION_UNIT(ctCuring.getOPERATION_UNIT());
                currentCTCuring.setBASE_QUANTITY(ctCuring.getBASE_QUANTITY());
                currentCTCuring.setSTANDART_VALUE_UNIT(ctCuring.getSTANDART_VALUE_UNIT());
                currentCTCuring.setCT_SEC1(ctCuring.getCT_SEC1());
                currentCTCuring.setCT_HR1000(ctCuring.getCT_HR1000());
                currentCTCuring.setWH_NORMAL_SHIFT_0(ctCuring.getWH_NORMAL_SHIFT_0());
                currentCTCuring.setWH_NORMAL_SHIFT_1(ctCuring.getWH_NORMAL_SHIFT_1());
                currentCTCuring.setWH_NORMAL_SHIFT_2(ctCuring.getWH_NORMAL_SHIFT_2());
                currentCTCuring.setWH_SHIFT_FRIDAY(ctCuring.getWH_SHIFT_FRIDAY()); 
                currentCTCuring.setWH_TOTAL_NORMAL_SHIFT(ctCuring.getWH_TOTAL_NORMAL_SHIFT());
                currentCTCuring.setWH_TOTAL_SHIFT_FRIDAY(ctCuring.getWH_TOTAL_SHIFT_FRIDAY()); 
                currentCTCuring.setALLOW_NORMAL_SHIFT_0(ctCuring.getALLOW_NORMAL_SHIFT_0());
                currentCTCuring.setALLOW_NORMAL_SHIFT_1(ctCuring.getALLOW_NORMAL_SHIFT_1());
                currentCTCuring.setALLOW_NORMAL_SHIFT_2(ctCuring.getALLOW_NORMAL_SHIFT_2());
                currentCTCuring.setALLOW_TOTAL(ctCuring.getALLOW_TOTAL());
                currentCTCuring.setOP_TIME_NORMAL_SHIFT_0(ctCuring.getOP_TIME_NORMAL_SHIFT_0());
                currentCTCuring.setOP_TIME_NORMAL_SHIFT_1(ctCuring.getOP_TIME_NORMAL_SHIFT_1());
                currentCTCuring.setOP_TIME_NORMAL_SHIFT_2(ctCuring.getOP_TIME_NORMAL_SHIFT_2());
                currentCTCuring.setOP_TIME_SHIFT_FRIDAY(ctCuring.getOP_TIME_SHIFT_FRIDAY());  
                currentCTCuring.setOP_TIME_NORMAL_SHIFT(ctCuring.getOP_TIME_NORMAL_SHIFT());
                currentCTCuring.setOP_TIME_TOTAL_SHIFT_FRIDAY(ctCuring.getOP_TIME_TOTAL_SHIFT_FRIDAY()); 
                currentCTCuring.setKAPS_NORMAL_SHIFT_0(ctCuring.getKAPS_NORMAL_SHIFT_0());
                currentCTCuring.setKAPS_NORMAL_SHIFT_1(ctCuring.getKAPS_NORMAL_SHIFT_1());
                currentCTCuring.setKAPS_NORMAL_SHIFT_2(ctCuring.getKAPS_NORMAL_SHIFT_2());
                currentCTCuring.setKAPS_SHIFT_FRIDAY(ctCuring.getKAPS_SHIFT_FRIDAY());
                currentCTCuring.setKAPS_TOTAL_NORMAL_SHIFT(ctCuring.getKAPS_TOTAL_NORMAL_SHIFT());
                currentCTCuring.setKAPS_TOTAL_SHIFT_FRIDAY(ctCuring.getKAPS_TOTAL_SHIFT_FRIDAY());  
                currentCTCuring.setWAKTU_TOTAL_CT_NORMAL(ctCuring.getWAKTU_TOTAL_CT_NORMAL());
                currentCTCuring.setWAKTU_TOTAL_CT_FRIDAY(ctCuring.getWAKTU_TOTAL_CT_FRIDAY());  
                
                currentCTCuring.setSTATUS(ctCuring.getSTATUS()); 
                currentCTCuring.setLAST_UPDATE_DATE(new Date());  
                currentCTCuring.setLAST_UPDATED_BY(ctCuring.getLAST_UPDATED_BY());
                
                return ctCuringRepo.save(currentCTCuring);
            } else {
                throw new RuntimeException("CT CURING with ID " + ctCuring.getCT_CURING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating CT CURING: " + e.getMessage());
            throw e;
        }
    }

    
    public CTCuring deleteCTCuring(CTCuring ctCuring) {
        try {
            Optional<CTCuring> currentCTCuringOpt = ctCuringRepo.findById(ctCuring.getCT_CURING_ID());
            
            if (currentCTCuringOpt.isPresent()) {
            	CTCuring currentCTCuring = currentCTCuringOpt.get();
                
            	currentCTCuring.setSTATUS(BigDecimal.valueOf(0));
            	currentCTCuring.setLAST_UPDATE_DATE(new Date());
            	currentCTCuring.setLAST_UPDATED_BY(ctCuring.getLAST_UPDATED_BY());
                
                return ctCuringRepo.save(currentCTCuring);
            } else {
                throw new RuntimeException("CY CURING with ID " + ctCuring.getCT_CURING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating CT CURING: " + e.getMessage());
            throw e;
        }
    }
    public CTCuring activateCTCuring(CTCuring ctCuring) {
        try {
            Optional<CTCuring> currentCTCuringOpt = ctCuringRepo.findById(ctCuring.getCT_CURING_ID());
            
            if (currentCTCuringOpt.isPresent()) {
            	CTCuring currentCTCuring = currentCTCuringOpt.get();
                
            	currentCTCuring.setSTATUS(BigDecimal.valueOf(1));
            	currentCTCuring.setLAST_UPDATE_DATE(new Date());
            	currentCTCuring.setLAST_UPDATED_BY(ctCuring.getLAST_UPDATED_BY());
                
                return ctCuringRepo.save(currentCTCuring);
            } else {
                throw new RuntimeException("CY CURING with ID " + ctCuring.getCT_CURING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating CT CURING: " + e.getMessage());
            throw e;
        }
    }
    
    
    public void deleteAllCTCuring() {
    	ctCuringRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportCTCuringsExcel() throws IOException {
        List<CTCuring> ctCurings = ctCuringRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(ctCurings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<CTCuring> ctCurings) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
                List<MachineCuring> activeMachineCurings = machineCuringRepo.findMachineCuringActive();
            List<ItemCuring> activeItemCurings = itemCuringRepo.findItemCuringActive();
            List<String> machineCuringWCT = activeMachineCurings.stream()
                .map(MachineCuring::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());
             List<String> itemCuringID = activeItemCurings.stream()
                 .map(ItemCuring::getITEM_CURING)
                 .collect(Collectors.toList());
            
            Sheet sheet = workbook.createSheet("CT_CURING DATA");

            // Create fonts for header and sub-header
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setColor(IndexedColors.BLACK.getIndex());

            Font subHeaderFont = workbook.createFont();
            subHeaderFont.setBold(true);
            subHeaderFont.setFontHeightInPoints((short) 11);
            
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            
            byte[] whiteColor = new byte[] {(byte) 255, (byte) 255, (byte) 255};
            byte[] greyColor = new byte[] {(byte) 191, (byte) 191, (byte) 191};
            byte[] brownColor = new byte[] {(byte) 191, (byte) 143, (byte) 0};
            byte[] babyblueColor = new byte[] {(byte) 142, (byte) 169, (byte) 219};
            byte[] yellowColor = new byte[] {(byte) 255, (byte) 255, (byte) 0};
            byte[] sageGreenColor = new byte[] {(byte) 169, (byte) 208, (byte) 142};
            byte[] orangeColor = new byte[] {(byte) 244, (byte) 176, (byte) 132};
            byte[] greenColor = new byte[] {(byte) 0, (byte) 176, (byte) 80};
            byte[] leafgreenColor = new byte[] {(byte) 146, (byte) 208, (byte) 80};
            byte[] blueColor = new byte[] {(byte) 0, (byte) 112, (byte) 192};
            byte[] skyblueColor = new byte[] {(byte) 0, (byte) 176, (byte) 240};


        // Create cell styles with background colors
            CellStyle whiteStyle = createCellStyleWithBorder(workbook, whiteColor);
            CellStyle lightGreyStyle = createCellStyleWithBorder(workbook, greyColor);
            CellStyle lightBrownStyle = createCellStyleWithBorder(workbook, brownColor);
            CellStyle skyBlueStyle = createCellStyleWithBorder(workbook, babyblueColor);
            CellStyle yellowStyle = createCellStyleWithBorder(workbook, yellowColor);
            CellStyle darkGreenStyle = createCellStyleWithBorder(workbook, greenColor);
            CellStyle lightGreenStyle = createCellStyleWithBorder(workbook, leafgreenColor);
            CellStyle sageGreenStyle = createCellStyleWithBorder(workbook, sageGreenColor); 
            CellStyle lightOrangeStyle = createCellStyleWithBorder(workbook, orangeColor); 
            CellStyle darkBlueStyle = createCellStyleWithBorder(workbook, blueColor);
            CellStyle lightBlueStyle = createCellStyleWithBorder(workbook, skyblueColor);
        
            // Create the first header row and merge cells
            Row headerRow0 = sheet.createRow(0);
            createMergedHeaderCell(sheet, headerRow0, 0, 10, "BASE DATA ROUTING SAP", whiteStyle);
            createMergedHeaderCell(sheet, headerRow0, 11, 12, "JAM KERJA", whiteStyle);
            createHeaderCell(headerRow0, 13, ":", whiteStyle);
            createHeaderCell(headerRow0, 14, "", yellowStyle);
            createHeaderCell(headerRow0, 15, "BTOL", yellowStyle);
            createHeaderCell(headerRow0, 16, "", whiteStyle);
            createMergedHeaderCell(sheet, headerRow0, 17, 20, "ALLOWANCE", lightOrangeStyle);
            createMergedHeaderCell(sheet, headerRow0, 21, 26, "OPERATIONAL TIME", darkGreenStyle);
            createMergedHeaderCell(sheet, headerRow0, 27, 34, "KAPASITAS", darkBlueStyle);

            Row headerRow1 = sheet.createRow(1); 
            Row headerRow2 = sheet.createRow(2); 
            
            createMergedHeaderCell(sheet, 1, 2, 0, 0, "WIP", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 1, 1, "Group Counter", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 2, 2, "Var Group Counter", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 3, 3, "Sequence", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 4, 4, "WCT", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 5, 5, "Operation Short Text", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 6, 6, "Operation Unit", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 7, 7, "Base Quantity", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 8, 8, "Standard Value Unit", lightGreyStyle);

            
            createHeaderCell(headerRow1, 9, "CT", lightBrownStyle);
            createHeaderCell(headerRow1, 10, "CT", lightBrownStyle);
            createMergedHeaderCell(sheet, headerRow1, 11, 13, "normal shift-", skyBlueStyle);
            createHeaderCell(headerRow1, 14, "FRIDAY", sageGreenStyle);
            createMergedHeaderCell(sheet, 1, 2, 15, 15, "TOTAL NORMAL", skyBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 16, 16, "TOTAL FRIDAY", sageGreenStyle);
            createHeaderCell(headerRow1, 17, "NORMAL", lightOrangeStyle);
            createHeaderCell(headerRow1, 18, "NORMAL", lightOrangeStyle);
            createHeaderCell(headerRow1, 19, "NORMAL", lightOrangeStyle);
            createMergedHeaderCell(sheet, 1, 2, 20, 20, "TOTAL", lightOrangeStyle);
            createHeaderCell(headerRow1, 21, "NORMAL", lightGreenStyle);
            createHeaderCell(headerRow1, 22, "NORMAL", lightGreenStyle);
            createHeaderCell(headerRow1, 23, "NORMAL", lightGreenStyle);
            createHeaderCell(headerRow1, 24, "JUMAT", darkGreenStyle);
            createMergedHeaderCell(sheet, 1, 2, 25, 25, "TOTAL NORMAL", lightGreenStyle);
            createMergedHeaderCell(sheet, 1, 2, 26, 26, "TOTAL FRIDAY", darkGreenStyle);
            createHeaderCell(headerRow1, 27, "NORMAL", lightBlueStyle);
            createHeaderCell(headerRow1, 28, "NORMAL", lightBlueStyle);
            createHeaderCell(headerRow1, 29, "NORMAL", lightBlueStyle);
            createHeaderCell(headerRow1, 30, "JUMAT", darkBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 31, 31, "TOTAL NORMAL", lightBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 32, 32, "TOTAL FRIDAY", darkBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 33, 33, "WAKTU TOTAL /CT", lightBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 34, 34, "WAKTU TOTAL /CT FRIDAY", darkBlueStyle);


            createHeaderCell(headerRow2, 9, "SEKON/1PC", lightBrownStyle);
            createHeaderCell(headerRow2, 10, "HR/1000PC", lightBrownStyle);
            createHeaderCell(headerRow2, 11, "1", skyBlueStyle);
            createHeaderCell(headerRow2, 12, "2", skyBlueStyle);
            createHeaderCell(headerRow2, 13, "3", skyBlueStyle);
            createHeaderCell(headerRow2, 14, "4", sageGreenStyle);
            createHeaderCell(headerRow2, 17, "SHIFT 1", lightOrangeStyle);
            createHeaderCell(headerRow2, 18, "SHIFT 2", lightOrangeStyle);
            createHeaderCell(headerRow2, 19, "SHIFT 3", lightOrangeStyle);
            createHeaderCell(headerRow2, 21, "SHIFT 1", lightGreenStyle);
            createHeaderCell(headerRow2, 22, "SHIFT 2", lightGreenStyle);
            createHeaderCell(headerRow2, 23, "SHIFT 3", lightGreenStyle);
            createHeaderCell(headerRow2, 24, "SHIFT 1", darkGreenStyle);
            createHeaderCell(headerRow2, 27, "SHIFT 1", lightBlueStyle);
            createHeaderCell(headerRow2, 28, "SHIFT 2", lightBlueStyle);
            createHeaderCell(headerRow2, 29, "SHIFT 3", lightBlueStyle);
            createHeaderCell(headerRow2, 30, "SHIFT 1", darkBlueStyle);


            // Populate data rows with consistent cell styling
            int rowIndex = 3; 
            for (CTCuring c : ctCurings) {
                Row dataRow = sheet.createRow(rowIndex++);

                // WIP
                dataRow.createCell(0).setCellValue(c.getWIP() != null ? c.getWIP() : "");

                // GROUP_COUNTER
                dataRow.createCell(1).setCellValue(c.getGROUP_COUNTER() != null ? c.getGROUP_COUNTER() : "");

                // VAR_GROUP_COUNTER
                dataRow.createCell(2).setCellValue(c.getVAR_GROUP_COUNTER() != null ? c.getVAR_GROUP_COUNTER() : "");

                // SEQUENCE
                dataRow.createCell(3).setCellValue(c.getSEQUENCE() != null ? c.getSEQUENCE().doubleValue() : null);

                // WCT
                dataRow.createCell(4).setCellValue(c.getWCT() != null ? c.getWCT() : "");

                // OPERATION_SHORT_TEXT
                dataRow.createCell(5).setCellValue(c.getOPERATION_SHORT_TEXT() != null ? c.getOPERATION_SHORT_TEXT() : "");

                // OPERATION_UNIT
                dataRow.createCell(6).setCellValue(c.getOPERATION_UNIT() != null ? c.getOPERATION_UNIT() : "");

                // BASE_QUANTITY
                dataRow.createCell(7).setCellValue(c.getBASE_QUANTITY() != null ? c.getBASE_QUANTITY().doubleValue() : null);

                // STANDARD_VALUE_UNIT
                dataRow.createCell(8).setCellValue(c.getSTANDART_VALUE_UNIT() != null ? c.getSTANDART_VALUE_UNIT() : "");

                // CT_SEC_1
                dataRow.createCell(9).setCellValue(c.getCT_SEC1() != null ? c.getCT_SEC1().doubleValue() : null);

                // CT_HR_1000
                dataRow.createCell(10).setCellValue(c.getCT_HR1000() != null ? c.getCT_HR1000().doubleValue() : null);

                // WH_NORMAL_SHIFT_0
                dataRow.createCell(11).setCellValue(c.getWH_NORMAL_SHIFT_0() != null ? c.getWH_NORMAL_SHIFT_0().doubleValue() : null);

                // WH_NORMAL_SHIFT_1
                dataRow.createCell(12).setCellValue(c.getWH_NORMAL_SHIFT_1() != null ? c.getWH_NORMAL_SHIFT_1().doubleValue() : null);

                // WH_NORMAL_SHIFT_2
                dataRow.createCell(13).setCellValue(c.getWH_NORMAL_SHIFT_2() != null ? c.getWH_NORMAL_SHIFT_2().doubleValue() : null);

                // WH_SHIFT_FRIDAY
                dataRow.createCell(14).setCellValue(c.getWH_SHIFT_FRIDAY() != null ? c.getWH_SHIFT_FRIDAY().doubleValue() : null);

                // WH_TOTAL_NORMAL_SHIFT
                dataRow.createCell(15).setCellValue(c.getWH_TOTAL_NORMAL_SHIFT() != null ? c.getWH_TOTAL_NORMAL_SHIFT().doubleValue() : null);

                // WH_TOTAL_SHIFT_FRIDAY
                dataRow.createCell(16).setCellValue(c.getWH_TOTAL_SHIFT_FRIDAY() != null ? c.getWH_TOTAL_SHIFT_FRIDAY().doubleValue() : null);

                // ALLOW_NORMAL_SHIFT_0
                dataRow.createCell(17).setCellValue(c.getALLOW_NORMAL_SHIFT_0() != null ? c.getALLOW_NORMAL_SHIFT_0().doubleValue() : null);

                // ALLOW_NORMAL_SHIFT_1
                dataRow.createCell(18).setCellValue(c.getALLOW_NORMAL_SHIFT_1() != null ? c.getALLOW_NORMAL_SHIFT_1().doubleValue() : null);

                // ALLOW_NORMAL_SHIFT_2
                dataRow.createCell(19).setCellValue(c.getALLOW_NORMAL_SHIFT_2() != null ? c.getALLOW_NORMAL_SHIFT_2().doubleValue() : null);

                // ALLOW_TOTAL
                dataRow.createCell(20).setCellValue(c.getALLOW_TOTAL() != null ? c.getALLOW_TOTAL().doubleValue() : null);

                // OP_TIME_NORMAL_SHIFT_0
                dataRow.createCell(21).setCellValue(c.getOP_TIME_NORMAL_SHIFT_0() != null ? c.getOP_TIME_NORMAL_SHIFT_0().doubleValue() : null);

                // OP_TIME_NORMAL_SHIFT_1
                dataRow.createCell(22).setCellValue(c.getOP_TIME_NORMAL_SHIFT_1() != null ? c.getOP_TIME_NORMAL_SHIFT_1().doubleValue() : null);

                // OP_TIME_NORMAL_SHIFT_2
                dataRow.createCell(23).setCellValue(c.getOP_TIME_NORMAL_SHIFT_2() != null ? c.getOP_TIME_NORMAL_SHIFT_2().doubleValue() : null);

                // OP_TIME_SHIFT_FRIDAY
                dataRow.createCell(24).setCellValue(c.getOP_TIME_SHIFT_FRIDAY() != null ? c.getOP_TIME_SHIFT_FRIDAY().doubleValue() : null);

                // OP_TIME_TOTAL_NORMAL_SHIFT
                dataRow.createCell(25).setCellValue(c.getOP_TIME_NORMAL_SHIFT() != null ? c.getOP_TIME_NORMAL_SHIFT().doubleValue() : null);

                // OP_TIME_TOTAL_SHIFT_FRIDAY
                dataRow.createCell(26).setCellValue(c.getOP_TIME_TOTAL_SHIFT_FRIDAY() != null ? c.getOP_TIME_TOTAL_SHIFT_FRIDAY().doubleValue() : null);

                // KAPS_NORMAL_SHIFT_0
                dataRow.createCell(27).setCellValue(c.getKAPS_NORMAL_SHIFT_0() != null ? c.getKAPS_NORMAL_SHIFT_0().doubleValue() : null);

                // KAPS_NORMAL_SHIFT_1
                dataRow.createCell(28).setCellValue(c.getKAPS_NORMAL_SHIFT_1() != null ? c.getKAPS_NORMAL_SHIFT_1().doubleValue() : null);

                // KAPS_NORMAL_SHIFT_2
                dataRow.createCell(29).setCellValue(c.getKAPS_NORMAL_SHIFT_2() != null ? c.getKAPS_NORMAL_SHIFT_2().doubleValue() : null);

                // KAPS_SHIFT_FRIDAY
                dataRow.createCell(30).setCellValue(c.getKAPS_SHIFT_FRIDAY() != null ? c.getKAPS_SHIFT_FRIDAY().doubleValue() : null);

                // KAPS_TOTAL_NORMAL_SHIFT
                dataRow.createCell(31).setCellValue(c.getKAPS_TOTAL_NORMAL_SHIFT() != null ? c.getKAPS_TOTAL_NORMAL_SHIFT().doubleValue() : null);

                // KAPS_TOTAL_SHIFT_FRIDAY
                dataRow.createCell(32).setCellValue(c.getKAPS_TOTAL_SHIFT_FRIDAY() != null ? c.getKAPS_TOTAL_SHIFT_FRIDAY().doubleValue() : null);

                // WAKTU_TOTAL_CT_NORMAL
                dataRow.createCell(33).setCellValue(c.getWAKTU_TOTAL_CT_NORMAL() != null ? c.getWAKTU_TOTAL_CT_NORMAL().doubleValue() : null);

                // WAKTU_TOTAL_CT_FRIDAY
                dataRow.createCell(34).setCellValue(c.getWAKTU_TOTAL_CT_FRIDAY() != null ? c.getWAKTU_TOTAL_CT_FRIDAY().doubleValue() : null);

                // Apply border style to each cell in the data row
                for (int i = 0; i <= 34; i++) {
                    dataRow.getCell(i).setCellStyle(borderStyle);
                }
            }
            sheet.setColumnWidth(0, 20 * 256);
            sheet.setColumnWidth(1, 20 * 256);  // Kolom C: Group Counter
            sheet.setColumnWidth(2, 20 * 256);  // Kolom D: Var Group Counter
            sheet.setColumnWidth(3, 20 * 256);  // Kolom E: Sequence
            sheet.setColumnWidth(4, 10 * 256);  // Kolom F: WCT
            sheet.setColumnWidth(5, 35 * 256);  // Kolom G: Operation Short Text
            sheet.setColumnWidth(6, 30 * 256);  // Kolom H: Operation Unit
            sheet.setColumnWidth(7, 15 * 256);  // Kolom I: Base Quantity
            sheet.setColumnWidth(8, 20 * 256);  // Kolom J: Standard Value Unit
            sheet.setColumnWidth(9, 15 * 256); // Kolom K: CT_SEC_1
            sheet.setColumnWidth(10, 15 * 256); // Kolom L: CT_HR_1000
            sheet.setColumnWidth(11, 12 * 256); // Kolom M: WH_NORMAL_SHIFT_0
            sheet.setColumnWidth(12, 12 * 256); // Kolom N: WH_NORMAL_SHIFT_1
            sheet.setColumnWidth(13, 12 * 256); // Kolom O: WH_NORMAL_SHIFT_2
            sheet.setColumnWidth(14, 12 * 256); // Kolom P: WH_SHIFT_FRIDAY
            sheet.setColumnWidth(15, 15 * 256); // Kolom Q: WH_TOTAL_NORMAL_SHIFT
            sheet.setColumnWidth(16, 15 * 256); // Kolom R: WH_TOTAL_SHIFT_FRIDAY
            sheet.setColumnWidth(17, 15 * 256); // Kolom S: ALLOW_NORMAL_SHIFT_0
            sheet.setColumnWidth(18, 15 * 256); // Kolom T: ALLOW_NORMAL_SHIFT_1
            sheet.setColumnWidth(19, 15 * 256); // Kolom U: ALLOW_NORMAL_SHIFT_2
            sheet.setColumnWidth(20, 15 * 256); // Kolom V: ALLOW_TOTAL
            sheet.setColumnWidth(21, 15 * 256); // Kolom W: OP_TIME_NORMAL_SHIFT_0
            sheet.setColumnWidth(22, 15 * 256); // Kolom X: OP_TIME_NORMAL_SHIFT_1
            sheet.setColumnWidth(23, 15 * 256); // Kolom Y: OP_TIME_NORMAL_SHIFT_2
            sheet.setColumnWidth(24, 15 * 256); // Kolom Z: OP_TIME_SHIFT_FRIDAY
            sheet.setColumnWidth(25, 15 * 256); // Kolom AA: OP_TIME_TOTAL_NORMAL_SHIFT
            sheet.setColumnWidth(26, 15 * 256); // Kolom AB: OP_TIME_TOTAL_SHIFT_FRIDAY
            sheet.setColumnWidth(27, 15 * 256); // Kolom AC: KAPS_NORMAL_SHIFT_0
            sheet.setColumnWidth(28, 15 * 256); // Kolom AD: KAPS_NORMAL_SHIFT_1
            sheet.setColumnWidth(29, 15 * 256); // Kolom AE: KAPS_NORMAL_SHIFT_2
            sheet.setColumnWidth(30, 15 * 256); // Kolom AF: KAPS_SHIFT_FRIDAY
            sheet.setColumnWidth(31, 15 * 256); // Kolom AG: KAPS_TOTAL_NORMAL_SHIFT
            sheet.setColumnWidth(32, 15 * 256); // Kolom AH: KAPS_TOTAL_SHIFT_FRIDAY
            sheet.setColumnWidth(33, 25 * 256); // Kolom AI: WAKTU_TOTAL_CT_NORMAL
            sheet.setColumnWidth(34, 25 * 256); // Kolom AJ: WAKTU_TOTAL_CT_FRIDAY

            Sheet hiddenSheetMachineCuring = workbook.createSheet("HIDDEN_MACHINECURINGS");
            for (int i = 0; i < machineCuringWCT.size(); i++) {
                Row row = hiddenSheetMachineCuring.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineCuringWCT.get(i));
            }

            Name namedRangeMachineCuring = workbook.createName();
            namedRangeMachineCuring.setNameName("machineCuringWCT");
            namedRangeMachineCuring.setRefersToFormula("HIDDEN_MACHINECURINGS!$A$1:$A$" + machineCuringWCT.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetMachineCuring), true);

            Sheet hiddenSheetItemCuring = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < itemCuringID.size(); i++) {
                Row row = hiddenSheetItemCuring.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(itemCuringID.get(i));
            }

            Name namedRangeItemCuring = workbook.createName();
            namedRangeItemCuring.setNameName("itemCuringID");
            namedRangeItemCuring.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetItemCuring), true);

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint itemCuringTypeConstraint = validationHelper.createFormulaListConstraint("itemCuringID");
            CellRangeAddressList itemCuringTypeAddressList = new CellRangeAddressList(3, 1000, 0, 0);
            DataValidation itemCuringTypeValidation = validationHelper.createValidation(itemCuringTypeConstraint, itemCuringTypeAddressList);
            itemCuringTypeValidation.setSuppressDropDownArrow(true);
            itemCuringTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(itemCuringTypeValidation);

            DataValidationConstraint machineCuringWCTConstraint = validationHelper.createFormulaListConstraint("machineCuringWCT");
            CellRangeAddressList machineCuringWCTAddressList = new CellRangeAddressList(3, 1000, 5, 5);
            DataValidation machineCuringWCTValidation = validationHelper.createValidation(machineCuringWCTConstraint, machineCuringWCTAddressList);
            machineCuringWCTValidation.setSuppressDropDownArrow(true);
            machineCuringWCTValidation.setShowErrorBox(true);
            sheet.addValidationData(machineCuringWCTValidation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export CT CURING data.");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
    
    private static CellStyle createCellStyleWithBorder(Workbook workbook, byte[] rgbColor) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor color = new XSSFColor(rgbColor, null);
        cellStyle.setFillForegroundColor(color);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        return cellStyle;
    }
    
    private void createMergedHeaderCell(Sheet sheet, Row row, int startCol, int endCol, String text, CellStyle style) {
        Cell cell = row.createCell(startCol);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        
        for (int col = startCol; col <= endCol; col++) {
            Cell currentCell = row.getCell(col);
            
            if (currentCell == null) {
                currentCell = row.createCell(col);
            }
            
            CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.cloneStyleFrom(style); 
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            
            cellStyle.setAlignment(HorizontalAlignment.CENTER); 
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            
            currentCell.setCellStyle(cellStyle);
        }
        
        sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), startCol, endCol));
    }

    private void createHeaderCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        
        CellStyle centeredStyle = row.getSheet().getWorkbook().createCellStyle();
        centeredStyle.cloneStyleFrom(style); 
        centeredStyle.setBorderTop(BorderStyle.THIN);
        centeredStyle.setBorderBottom(BorderStyle.THIN);
        centeredStyle.setBorderLeft(BorderStyle.THIN);
        centeredStyle.setBorderRight(BorderStyle.THIN);
        
        centeredStyle.setAlignment(HorizontalAlignment.CENTER);
        centeredStyle.setVerticalAlignment(VerticalAlignment.CENTER); 
        
        Sheet sheet = row.getSheet();
        
        cell.setCellStyle(centeredStyle);
    }

    private void createMergedHeaderCell(Sheet sheet, int startRow, int endRow, int startCol, int endCol, String text, CellStyle style) {
        Row row = sheet.getRow(startRow);
        if (row == null) {
            row = sheet.createRow(startRow);
        }

        Cell cell = row.createCell(startCol);
        cell.setCellValue(text);
        
        CellStyle mergedCellStyle = sheet.getWorkbook().createCellStyle();
        mergedCellStyle.cloneStyleFrom(style); 
        
        mergedCellStyle.setBorderTop(BorderStyle.THIN);
        mergedCellStyle.setBorderBottom(BorderStyle.THIN);
        mergedCellStyle.setBorderLeft(BorderStyle.THIN);
        mergedCellStyle.setBorderRight(BorderStyle.THIN);
        
        mergedCellStyle.setAlignment(HorizontalAlignment.CENTER);
        mergedCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); 
        
        cell.setCellStyle(mergedCellStyle);

        sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCol, endCol));
        
        for (int r = startRow; r <= endRow; r++) {
            Row mergedRow = sheet.getRow(r);
            if (mergedRow == null) {
                mergedRow = sheet.createRow(r);
            }
            for (int c = startCol; c <= endCol; c++) {
                Cell mergedCell = mergedRow.getCell(c);
                if (mergedCell == null) {
                    mergedCell = mergedRow.createCell(c);
                }
                mergedCell.setCellStyle(mergedCellStyle); 
            }
        }
    }
     public ByteArrayInputStream layoutCTCuringsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }

    private ByteArrayInputStream layoutToExcel() throws IOException {

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            List<MachineCuring> activeMachineCurings = machineCuringRepo.findMachineCuringActive();
            List<ItemCuring> activeItemCurings = itemCuringRepo.findItemCuringActive();
            List<String> machineCuringWCT = activeMachineCurings.stream()
                .map(MachineCuring::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());
             List<String> itemCuringID = activeItemCurings.stream()
                 .map(ItemCuring::getITEM_CURING)
                 .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("CT_CURING DATA");

            // Create fonts for header and sub-header
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setColor(IndexedColors.BLACK.getIndex());

            Font subHeaderFont = workbook.createFont();
            subHeaderFont.setBold(true);
            subHeaderFont.setFontHeightInPoints((short) 11);
            
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            
            byte[] whiteColor = new byte[] {(byte) 255, (byte) 255, (byte) 255};
            byte[] greyColor = new byte[] {(byte) 191, (byte) 191, (byte) 191};
            byte[] brownColor = new byte[] {(byte) 191, (byte) 143, (byte) 0};
            byte[] babyblueColor = new byte[] {(byte) 142, (byte) 169, (byte) 219};
            byte[] yellowColor = new byte[] {(byte) 255, (byte) 255, (byte) 0};
            byte[] sageGreenColor = new byte[] {(byte) 169, (byte) 208, (byte) 142};
            byte[] orangeColor = new byte[] {(byte) 244, (byte) 176, (byte) 132};
            byte[] greenColor = new byte[] {(byte) 0, (byte) 176, (byte) 80};
            byte[] leafgreenColor = new byte[] {(byte) 146, (byte) 208, (byte) 80};
            byte[] blueColor = new byte[] {(byte) 0, (byte) 112, (byte) 192};
            byte[] skyblueColor = new byte[] {(byte) 0, (byte) 176, (byte) 240};


        // Create cell styles with background colors
            CellStyle whiteStyle = createCellStyleWithBorder(workbook, whiteColor);
            CellStyle lightGreyStyle = createCellStyleWithBorder(workbook, greyColor);
            CellStyle lightBrownStyle = createCellStyleWithBorder(workbook, brownColor);
            CellStyle skyBlueStyle = createCellStyleWithBorder(workbook, babyblueColor);
            CellStyle yellowStyle = createCellStyleWithBorder(workbook, yellowColor);
            CellStyle darkGreenStyle = createCellStyleWithBorder(workbook, greenColor);
            CellStyle lightGreenStyle = createCellStyleWithBorder(workbook, leafgreenColor);
            CellStyle sageGreenStyle = createCellStyleWithBorder(workbook, sageGreenColor); 
            CellStyle lightOrangeStyle = createCellStyleWithBorder(workbook, orangeColor); 
            CellStyle darkBlueStyle = createCellStyleWithBorder(workbook, blueColor);
            CellStyle lightBlueStyle = createCellStyleWithBorder(workbook, skyblueColor);
        
            // Create the first header row and merge cells
            Row headerRow0 = sheet.createRow(0);
            createMergedHeaderCell(sheet, headerRow0, 0, 10, "BASE DATA ROUTING SAP", whiteStyle);
            createMergedHeaderCell(sheet, headerRow0, 11, 12, "JAM KERJA", whiteStyle);
            createHeaderCell(headerRow0, 13, ":", whiteStyle);
            createHeaderCell(headerRow0, 14, "", yellowStyle);
            createHeaderCell(headerRow0, 15, "BTOL", yellowStyle);
            createHeaderCell(headerRow0, 16, "", whiteStyle);
            createMergedHeaderCell(sheet, headerRow0, 17, 20, "ALLOWANCE", lightOrangeStyle);
            createMergedHeaderCell(sheet, headerRow0, 21, 26, "OPERATIONAL TIME", darkGreenStyle);
            createMergedHeaderCell(sheet, headerRow0, 27, 34, "KAPASITAS", darkBlueStyle);

            Row headerRow1 = sheet.createRow(1); 
            Row headerRow2 = sheet.createRow(2); 
            
            createMergedHeaderCell(sheet, 1, 2, 0, 0, "WIP", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 1, 1, "Group Counter", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 2, 2, "Var Group Counter", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 3, 3, "Sequence", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 4, 4, "WCT", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 5, 5, "Operation Short Text", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 6, 6, "Operation Unit", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 7, 7, "Base Quantity", lightGreyStyle);
            createMergedHeaderCell(sheet, 1, 2, 8, 8, "Standard Value Unit", lightGreyStyle);

            
            createHeaderCell(headerRow1, 9, "CT", lightBrownStyle);
            createHeaderCell(headerRow1, 10, "CT", lightBrownStyle);
            createMergedHeaderCell(sheet, headerRow1, 11, 13, "normal shift-", skyBlueStyle);
            createHeaderCell(headerRow1, 14, "FRIDAY", sageGreenStyle);
            createMergedHeaderCell(sheet, 1, 2, 15, 15, "TOTAL NORMAL", skyBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 16, 16, "TOTAL FRIDAY", sageGreenStyle);
            createHeaderCell(headerRow1, 17, "NORMAL", lightOrangeStyle);
            createHeaderCell(headerRow1, 18, "NORMAL", lightOrangeStyle);
            createHeaderCell(headerRow1, 19, "NORMAL", lightOrangeStyle);
            createMergedHeaderCell(sheet, 1, 2, 20, 20, "TOTAL", lightOrangeStyle);
            createHeaderCell(headerRow1, 21, "NORMAL", lightGreenStyle);
            createHeaderCell(headerRow1, 22, "NORMAL", lightGreenStyle);
            createHeaderCell(headerRow1, 23, "NORMAL", lightGreenStyle);
            createHeaderCell(headerRow1, 24, "JUMAT", darkGreenStyle);
            createMergedHeaderCell(sheet, 1, 2, 25, 25, "TOTAL NORMAL", lightGreenStyle);
            createMergedHeaderCell(sheet, 1, 2, 26, 26, "TOTAL FRIDAY", darkGreenStyle);
            createHeaderCell(headerRow1, 27, "NORMAL", lightBlueStyle);
            createHeaderCell(headerRow1, 28, "NORMAL", lightBlueStyle);
            createHeaderCell(headerRow1, 29, "NORMAL", lightBlueStyle);
            createHeaderCell(headerRow1, 30, "JUMAT", darkBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 31, 31, "TOTAL NORMAL", lightBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 32, 32, "TOTAL FRIDAY", darkBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 33, 33, "WAKTU TOTAL /CT", lightBlueStyle);
            createMergedHeaderCell(sheet, 1, 2, 34, 34, "WAKTU TOTAL /CT FRIDAY", darkBlueStyle);


            createHeaderCell(headerRow2, 9, "SEKON/1PC", lightBrownStyle);
            createHeaderCell(headerRow2, 10, "HR/1000PC", lightBrownStyle);
            createHeaderCell(headerRow2, 11, "1", skyBlueStyle);
            createHeaderCell(headerRow2, 12, "2", skyBlueStyle);
            createHeaderCell(headerRow2, 13, "3", skyBlueStyle);
            createHeaderCell(headerRow2, 14, "4", sageGreenStyle);
            createHeaderCell(headerRow2, 17, "SHIFT 1", lightOrangeStyle);
            createHeaderCell(headerRow2, 18, "SHIFT 2", lightOrangeStyle);
            createHeaderCell(headerRow2, 19, "SHIFT 3", lightOrangeStyle);
            createHeaderCell(headerRow2, 21, "SHIFT 1", lightGreenStyle);
            createHeaderCell(headerRow2, 22, "SHIFT 2", lightGreenStyle);
            createHeaderCell(headerRow2, 23, "SHIFT 3", lightGreenStyle);
            createHeaderCell(headerRow2, 24, "SHIFT 1", darkGreenStyle);
            createHeaderCell(headerRow2, 27, "SHIFT 1", lightBlueStyle);
            createHeaderCell(headerRow2, 28, "SHIFT 2", lightBlueStyle);
            createHeaderCell(headerRow2, 29, "SHIFT 3", lightBlueStyle);
            createHeaderCell(headerRow2, 30, "SHIFT 1", darkBlueStyle);
            sheet.setColumnWidth(0, 20 * 256);
            sheet.setColumnWidth(1, 20 * 256);  // Kolom C: Group Counter
            sheet.setColumnWidth(2, 20 * 256);  // Kolom D: Var Group Counter
            sheet.setColumnWidth(3, 20 * 256);  // Kolom E: Sequence
            sheet.setColumnWidth(4, 10 * 256);  // Kolom F: WCT
            sheet.setColumnWidth(5, 35 * 256);  // Kolom G: Operation Short Text
            sheet.setColumnWidth(6, 30 * 256);  // Kolom H: Operation Unit
            sheet.setColumnWidth(7, 15 * 256);  // Kolom I: Base Quantity
            sheet.setColumnWidth(8, 20 * 256);  // Kolom J: Standard Value Unit
            sheet.setColumnWidth(9, 15 * 256); // Kolom K: CT_SEC_1
            sheet.setColumnWidth(10, 15 * 256); // Kolom L: CT_HR_1000
            sheet.setColumnWidth(11, 12 * 256); // Kolom M: WH_NORMAL_SHIFT_0
            sheet.setColumnWidth(12, 12 * 256); // Kolom N: WH_NORMAL_SHIFT_1
            sheet.setColumnWidth(13, 12 * 256); // Kolom O: WH_NORMAL_SHIFT_2
            sheet.setColumnWidth(14, 12 * 256); // Kolom P: WH_SHIFT_FRIDAY
            sheet.setColumnWidth(15, 15 * 256); // Kolom Q: WH_TOTAL_NORMAL_SHIFT
            sheet.setColumnWidth(16, 15 * 256); // Kolom R: WH_TOTAL_SHIFT_FRIDAY
            sheet.setColumnWidth(17, 15 * 256); // Kolom S: ALLOW_NORMAL_SHIFT_0
            sheet.setColumnWidth(18, 15 * 256); // Kolom T: ALLOW_NORMAL_SHIFT_1
            sheet.setColumnWidth(19, 15 * 256); // Kolom U: ALLOW_NORMAL_SHIFT_2
            sheet.setColumnWidth(20, 15 * 256); // Kolom V: ALLOW_TOTAL
            sheet.setColumnWidth(21, 15 * 256); // Kolom W: OP_TIME_NORMAL_SHIFT_0
            sheet.setColumnWidth(22, 15 * 256); // Kolom X: OP_TIME_NORMAL_SHIFT_1
            sheet.setColumnWidth(23, 15 * 256); // Kolom Y: OP_TIME_NORMAL_SHIFT_2
            sheet.setColumnWidth(24, 15 * 256); // Kolom Z: OP_TIME_SHIFT_FRIDAY
            sheet.setColumnWidth(25, 15 * 256); // Kolom AA: OP_TIME_TOTAL_NORMAL_SHIFT
            sheet.setColumnWidth(26, 15 * 256); // Kolom AB: OP_TIME_TOTAL_SHIFT_FRIDAY
            sheet.setColumnWidth(27, 15 * 256); // Kolom AC: KAPS_NORMAL_SHIFT_0
            sheet.setColumnWidth(28, 15 * 256); // Kolom AD: KAPS_NORMAL_SHIFT_1
            sheet.setColumnWidth(29, 15 * 256); // Kolom AE: KAPS_NORMAL_SHIFT_2
            sheet.setColumnWidth(30, 15 * 256); // Kolom AF: KAPS_SHIFT_FRIDAY
            sheet.setColumnWidth(31, 15 * 256); // Kolom AG: KAPS_TOTAL_NORMAL_SHIFT
            sheet.setColumnWidth(32, 15 * 256); // Kolom AH: KAPS_TOTAL_SHIFT_FRIDAY
            sheet.setColumnWidth(33, 25 * 256); // Kolom AI: WAKTU_TOTAL_CT_NORMAL
            sheet.setColumnWidth(34, 25 * 256); // Kolom AJ: WAKTU_TOTAL_CT_FRIDAY
           

            Sheet hiddenSheetMachineCuring = workbook.createSheet("HIDDEN_MACHINECURINGS");
            for (int i = 0; i < machineCuringWCT.size(); i++) {
                Row row = hiddenSheetMachineCuring.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(machineCuringWCT.get(i));
            }

            Name namedRangeMachineCuring = workbook.createName();
            namedRangeMachineCuring.setNameName("machineCuringWCT");
            namedRangeMachineCuring.setRefersToFormula("HIDDEN_MACHINECURINGS!$A$1:$A$" + machineCuringWCT.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetMachineCuring), true);

            Sheet hiddenSheetItemCuring = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < itemCuringID.size(); i++) {
                Row row = hiddenSheetItemCuring.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(itemCuringID.get(i));
            }

            Name namedRangeItemCuring = workbook.createName();
            namedRangeItemCuring.setNameName("itemCuringID");
            namedRangeItemCuring.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetItemCuring), true);

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            DataValidationConstraint itemCuringTypeConstraint = validationHelper.createFormulaListConstraint("itemCuringID");
            CellRangeAddressList itemCuringTypeAddressList = new CellRangeAddressList(3, 1000, 0, 0);
            DataValidation itemCuringTypeValidation = validationHelper.createValidation(itemCuringTypeConstraint, itemCuringTypeAddressList);
            itemCuringTypeValidation.setSuppressDropDownArrow(true);
            itemCuringTypeValidation.setShowErrorBox(true);
            sheet.addValidationData(itemCuringTypeValidation);

            DataValidationConstraint machineCuringWCTConstraint = validationHelper.createFormulaListConstraint("machineCuringWCT");
            CellRangeAddressList machineCuringWCTAddressList = new CellRangeAddressList(3, 1000, 5, 5);
            DataValidation machineCuringWCTValidation = validationHelper.createValidation(machineCuringWCTConstraint, machineCuringWCTAddressList);
            machineCuringWCTValidation.setSuppressDropDownArrow(true);
            machineCuringWCTValidation.setShowErrorBox(true);
            sheet.addValidationData(machineCuringWCTValidation);

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
