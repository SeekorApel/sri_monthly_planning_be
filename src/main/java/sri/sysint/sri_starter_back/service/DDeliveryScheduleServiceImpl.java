package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.ZoneId;
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

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.DDeliverySchedule;
import sri.sysint.sri_starter_back.model.DeliverySchedule;
import sri.sysint.sri_starter_back.repository.DDeliveryScheduleRepo;
import sri.sysint.sri_starter_back.repository.DeliveryScheduleRepo;

@Service
@Transactional
public class DDeliveryScheduleServiceImpl {
    @Autowired
    private DDeliveryScheduleRepo dDeliveryScheduleRepo;

    @Autowired
    private DeliveryScheduleRepo deliveryScheduleRepo;

    public DDeliveryScheduleServiceImpl(DDeliveryScheduleRepo dDeliveryScheduleRepo) {
        this.dDeliveryScheduleRepo = dDeliveryScheduleRepo;
    }

    public BigDecimal getNewId() {
        return dDeliveryScheduleRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<DDeliverySchedule> getAllDDeliverySchedule() {
        Iterable<DDeliverySchedule> dDeliverySchedules = dDeliveryScheduleRepo.getDataOrderId();
        List<DDeliverySchedule> dDeliveryScheduleList = new ArrayList<>();
        for (DDeliverySchedule item : dDeliverySchedules) {
            DDeliverySchedule dDeliveryScheduleTemp = new DDeliverySchedule(item);
            dDeliveryScheduleList.add(dDeliveryScheduleTemp);
        }

        return dDeliveryScheduleList;
    }

    public Optional<DDeliverySchedule> getDDeliveryScheduleById(BigDecimal id) {
        Optional<DDeliverySchedule> dDeliverySchedule = dDeliveryScheduleRepo.findById(id);
        return dDeliverySchedule;
    }

    public DDeliverySchedule saveDDeliverySchedule(DDeliverySchedule dDeliverySchedule) {
        try {
            LocalDateTime date = dDeliverySchedule.getDATE_DS().toInstant()
                    .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
            
            dDeliverySchedule.setDATE_DS(Date.from(date.atZone(ZoneId.of("UTC")).toInstant()));
            
            dDeliverySchedule.setDETAIL_DS_ID(getNewId());
            dDeliverySchedule.setSTATUS(BigDecimal.valueOf(1));
            dDeliverySchedule.setCREATION_DATE(new Date());
            dDeliverySchedule.setLAST_UPDATE_DATE(new Date());
            return dDeliveryScheduleRepo.save(dDeliverySchedule);
        } catch (Exception e) {
            System.err.println("Error saving DDeliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    public DDeliverySchedule updateDDeliverySchedule(DDeliverySchedule dDeliverySchedule) {
        try {
            Optional<DDeliverySchedule> currentDDeliveryScheduleOpt = dDeliveryScheduleRepo.findById(dDeliverySchedule.getDETAIL_DS_ID());

            if (currentDDeliveryScheduleOpt.isPresent()) {
                DDeliverySchedule currentDDeliverySchedule = currentDDeliveryScheduleOpt.get();

                LocalDateTime date = dDeliverySchedule.getDATE_DS().toInstant()
                        .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
                
                currentDDeliverySchedule.setDATE_DS(Date.from(date.atZone(ZoneId.of("UTC")).toInstant()));
                
                currentDDeliverySchedule.setDS_ID(dDeliverySchedule.getDS_ID());
                currentDDeliverySchedule.setPART_NUM(dDeliverySchedule.getPART_NUM());
                currentDDeliverySchedule.setTOTAL_DELIVERY(dDeliverySchedule.getTOTAL_DELIVERY());
                currentDDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                currentDDeliverySchedule.setLAST_UPDATED_BY(dDeliverySchedule.getLAST_UPDATED_BY());

                return dDeliveryScheduleRepo.save(currentDDeliverySchedule);
            } else {
                throw new RuntimeException("DDeliverySchedule with ID " + dDeliverySchedule.getDETAIL_DS_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating DDeliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    public DDeliverySchedule deleteDDeliverySchedule(DDeliverySchedule dDeliverySchedule) {
        try {
            Optional<DDeliverySchedule> currentDDeliveryScheduleOpt = dDeliveryScheduleRepo.findById(dDeliverySchedule.getDETAIL_DS_ID());

            if (currentDDeliveryScheduleOpt.isPresent()) {
                DDeliverySchedule currentDDeliverySchedule = currentDDeliveryScheduleOpt.get();

                currentDDeliverySchedule.setSTATUS(BigDecimal.valueOf(0));
                currentDDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                currentDDeliverySchedule.setLAST_UPDATED_BY(dDeliverySchedule.getLAST_UPDATED_BY());

                return dDeliveryScheduleRepo.save(currentDDeliverySchedule);
            } else {
                throw new RuntimeException("DDeliverySchedule with ID " + dDeliverySchedule.getDETAIL_DS_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating DDeliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllDDeliverySchedule() {
        dDeliveryScheduleRepo.deleteAll();
    }
    
    public DDeliverySchedule restoreDDeliverySchedule(DDeliverySchedule dDeliverySchedule) {
        try {
            Optional<DDeliverySchedule> currentDDeliveryScheduleOpt = dDeliveryScheduleRepo.findById(dDeliverySchedule.getDETAIL_DS_ID());

            if (currentDDeliveryScheduleOpt.isPresent()) {
                DDeliverySchedule currentDDeliverySchedule = currentDDeliveryScheduleOpt.get();

                currentDDeliverySchedule.setSTATUS(BigDecimal.valueOf(1));
                currentDDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                currentDDeliverySchedule.setLAST_UPDATED_BY(dDeliverySchedule.getLAST_UPDATED_BY());

                return dDeliveryScheduleRepo.save(currentDDeliverySchedule);
            } else {
                throw new RuntimeException("DDeliverySchedule with ID " + dDeliverySchedule.getDETAIL_DS_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating DDeliverySchedule: " + e.getMessage());
            throw e;
        }
    }
    
    public ByteArrayInputStream exportDDeliverySchedulesExcel() throws IOException {
        List<DDeliverySchedule> dDeliverySchedules = dDeliveryScheduleRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(dDeliverySchedules);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<DDeliverySchedule> dDeliverySchedules) throws IOException {
        String[] header = {
            "NOMOR",                  
            "DETAIL_DS_ID",
            "DELIVERYSCHEDULE",
            "PART_NUM",
            "DATE",
            "TOTAL_DELIVERY"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("DDeliverySchedule Data");

            // Font for the header
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            
            // Border Style for cells
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            // Header Style with bold font and yellow background
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Date style for formatting dates in dd-MM-yyyy format
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.cloneStyleFrom(borderStyle);
            DataFormat dateFormat = workbook.createDataFormat();
            dateStyle.setDataFormat(dateFormat.getFormat("dd-MM-yyyy"));

            // Set column widths
            sheet.setColumnWidth(0, 10 * 256); // Nomor column
            sheet.setColumnWidth(1, 20 * 256); // DETAIL_DS_ID column
            sheet.setColumnWidth(2, 20 * 256); // DELIVERYSCHEDULE_ID column
            sheet.setColumnWidth(3, 30 * 256); // PART_NUM column
            sheet.setColumnWidth(4, 20 * 256); // DATE column
            sheet.setColumnWidth(5, 20 * 256); // TOTAL_DELIVERY column

            // Create the header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            // Fill data rows
            int rowIndex = 1;
            for (DDeliverySchedule d : dDeliverySchedules) {
                Row dataRow = sheet.createRow(rowIndex++);

                // Nomor column (sequential number)
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);  // Sequential number (1, 2, 3, ...)
                nomorCell.setCellStyle(borderStyle);

                // DETAIL_DS_ID column
                Cell detailDsIdCell = dataRow.createCell(1);
                detailDsIdCell.setCellValue(d.getDETAIL_DS_ID().doubleValue());
                detailDsIdCell.setCellStyle(borderStyle);

                // DELIVERYSCHEDULE_ID column
                Cell deliveryScheduleIdCell = dataRow.createCell(2);
                deliveryScheduleIdCell.setCellValue(d.getDS_ID() != null ? d.getDS_ID().doubleValue() : 0);
                deliveryScheduleIdCell.setCellStyle(borderStyle);

                // PART_NUM column
                Cell partNumCell = dataRow.createCell(3);
                partNumCell.setCellValue(d.getPART_NUM() != null ? d.getPART_NUM().doubleValue() : 0);
                partNumCell.setCellStyle(borderStyle);

                // DATE column (formatted as date)
                Cell dateCell = dataRow.createCell(4);
                if (d.getDATE_DS() != null) {
                    dateCell.setCellValue(d.getDATE_DS());
                    dateCell.setCellStyle(dateStyle);
                } else {
                    dateCell.setCellValue("");
                    dateCell.setCellStyle(borderStyle);
                }

                // TOTAL_DELIVERY column
                Cell totalDeliveryCell = dataRow.createCell(5);
                totalDeliveryCell.setCellValue(d.getTOTAL_DELIVERY() != null ? d.getTOTAL_DELIVERY().doubleValue() : 0);
                totalDeliveryCell.setCellStyle(borderStyle);
            }
            List<DeliverySchedule> activeDeliverySchedules = deliveryScheduleRepo.findDeliveryScheduleActive();
            List<BigDecimal> DeliveryScheduleIDs = activeDeliverySchedules.stream()
                .map(DeliverySchedule::getDS_ID)
                .collect(Collectors.toList());
                
            Sheet hiddenSheet = workbook.createSheet("HIDDEN_DELIVERYSCHEDULES");
            for (int i = 0; i < DeliveryScheduleIDs.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(DeliveryScheduleIDs.get(i).toPlainString());
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("DeliveryScheduleIDs");
            namedRange.setRefersToFormula("HIDDEN_DELIVERYSCHEDULES!$A$1:$A$" + DeliveryScheduleIDs.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("DeliveryScheduleIDs");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export DDeliverySchedule data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

        public ByteArrayInputStream layoutDDeliveryScheduleExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel( );
        return byteArrayInputStream;
    }
    
    public ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",                  
            "DETAIL_DS_ID",
            "DELIVERYSCHEDULE",
            "PART_NUM",
            "DATE",
            "TOTAL_DELIVERY"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<DeliverySchedule> activeDeliverySchedules = deliveryScheduleRepo.findDeliveryScheduleActive();
            List<BigDecimal> DeliveryScheduleIDs = activeDeliverySchedules.stream()
                .map(DeliverySchedule::getDS_ID)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("DDeliverySchedule Data");
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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_DELIVERYSCHEDULES");
            for (int i = 0; i < DeliveryScheduleIDs.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(DeliveryScheduleIDs.get(i).toPlainString());
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("DeliveryScheduleIDs");
            namedRange.setRefersToFormula("HIDDEN_DELIVERYSCHEDULES!$A$1:$A$" + DeliveryScheduleIDs.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            int nomor = 1;
            
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("DeliveryScheduleIDs");
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
