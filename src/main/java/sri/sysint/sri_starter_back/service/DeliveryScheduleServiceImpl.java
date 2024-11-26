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
import sri.sysint.sri_starter_back.model.DeliverySchedule;
import sri.sysint.sri_starter_back.repository.DeliveryScheduleRepo;

@Service
@Transactional
public class DeliveryScheduleServiceImpl {
    @Autowired
    private DeliveryScheduleRepo deliveryScheduleRepo;

    public DeliveryScheduleServiceImpl(DeliveryScheduleRepo deliveryScheduleRepo) {
        this.deliveryScheduleRepo = deliveryScheduleRepo;
    }

    public BigDecimal getNewId() {
        return deliveryScheduleRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<DeliverySchedule> getAllDeliverySchedules() {
        Iterable<DeliverySchedule> deliverySchedules = deliveryScheduleRepo.getDataOrderId();
        List<DeliverySchedule> deliveryScheduleList = new ArrayList<>();
        for (DeliverySchedule item : deliverySchedules) {
            DeliverySchedule deliveryScheduleTemp = new DeliverySchedule(item);
            deliveryScheduleList.add(deliveryScheduleTemp);
        }

        return deliveryScheduleList;
    }

    public Optional<DeliverySchedule> getDeliveryScheduleById(BigDecimal id) {
        Optional<DeliverySchedule> deliverySchedule = deliveryScheduleRepo.findById(id);
        return deliverySchedule;
    }

    public DeliverySchedule saveDeliverySchedule(DeliverySchedule deliverySchedule) {
        try {
            deliverySchedule.setDS_ID(getNewId());
            deliverySchedule.setSTATUS(BigDecimal.valueOf(1));
            deliverySchedule.setCREATION_DATE(new Date());
            deliverySchedule.setLAST_UPDATE_DATE(new Date());
            return deliveryScheduleRepo.save(deliverySchedule);
        } catch (Exception e) {
            System.err.println("Error saving deliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    public DeliverySchedule updateDeliverySchedule(DeliverySchedule deliverySchedule) {
        try {
            Optional<DeliverySchedule> currentDeliveryScheduleOpt = deliveryScheduleRepo.findById(deliverySchedule.getDS_ID());

            if (currentDeliveryScheduleOpt.isPresent()) {
                DeliverySchedule currentDeliverySchedule = currentDeliveryScheduleOpt.get();

                currentDeliverySchedule.setEFFECTIVE_TIME(deliverySchedule.getEFFECTIVE_TIME());
                currentDeliverySchedule.setDATE_ISSUED(deliverySchedule.getDATE_ISSUED());
                currentDeliverySchedule.setCATEGORY(deliverySchedule.getCATEGORY());
                currentDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                currentDeliverySchedule.setLAST_UPDATED_BY(deliverySchedule.getLAST_UPDATED_BY());

                return deliveryScheduleRepo.save(currentDeliverySchedule);
            } else {
                throw new RuntimeException("DeliverySchedule with ID " + deliverySchedule.getDS_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating deliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    public DeliverySchedule deleteDeliverySchedule(DeliverySchedule deliverySchedule) {
        try {
            Optional<DeliverySchedule> currentDeliveryScheduleOpt = deliveryScheduleRepo.findById(deliverySchedule.getDS_ID());

            if (currentDeliveryScheduleOpt.isPresent()) {
                DeliverySchedule currentDeliverySchedule = currentDeliveryScheduleOpt.get();

                currentDeliverySchedule.setSTATUS(BigDecimal.valueOf(0));
                currentDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                currentDeliverySchedule.setLAST_UPDATED_BY(deliverySchedule.getLAST_UPDATED_BY());

                return deliveryScheduleRepo.save(currentDeliverySchedule);
            } else {
                throw new RuntimeException("DeliverySchedule with ID " + deliverySchedule.getDS_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating deliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllDeliverySchedules() {
        deliveryScheduleRepo.deleteAll();
    }
    
    public DeliverySchedule restoreDeliverySchedule(DeliverySchedule deliverySchedule) {
        try {
            Optional<DeliverySchedule> currentDeliveryScheduleOpt = deliveryScheduleRepo.findById(deliverySchedule.getDS_ID());

            if (currentDeliveryScheduleOpt.isPresent()) {
                DeliverySchedule currentDeliverySchedule = currentDeliveryScheduleOpt.get();

                currentDeliverySchedule.setSTATUS(BigDecimal.valueOf(1));
                currentDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                currentDeliverySchedule.setLAST_UPDATED_BY(deliverySchedule.getLAST_UPDATED_BY());

                return deliveryScheduleRepo.save(currentDeliverySchedule);
            } else {
                throw new RuntimeException("DeliverySchedule with ID " + deliverySchedule.getDS_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating deliverySchedule: " + e.getMessage());
            throw e;
        }
    }

    
    public ByteArrayInputStream exportDeliverySchedulesExcel() throws IOException {
        List<DeliverySchedule> deliverySchedules = deliveryScheduleRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(deliverySchedules);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<DeliverySchedule> deliverySchedules) throws IOException {
        String[] header = {"NOMOR", "DELIVERYSCHEDULE_ID", "EFFECTIVE_TIME", "DATE_ISSUED", "CATEGORY"};

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("DELIVERY SCHEDULE DATA");

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
            for (DeliverySchedule ds : deliverySchedules) {
                Row dataRow = sheet.createRow(rowIndex);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(ds.getDS_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell effectiveTimeCell = dataRow.createCell(2);
                effectiveTimeCell.setCellValue(ds.getEFFECTIVE_TIME() != null ? ds.getEFFECTIVE_TIME().toString() : "");
                effectiveTimeCell.setCellStyle(borderStyle);

                Cell dateIssuedCell = dataRow.createCell(3);
                dateIssuedCell.setCellValue(ds.getDATE_ISSUED() != null ? ds.getDATE_ISSUED().toString() : "");
                dateIssuedCell.setCellStyle(borderStyle);

                Cell categoryCell = dataRow.createCell(4);
                categoryCell.setCellValue(ds.getCATEGORY() != null ? ds.getCATEGORY() : "");
                categoryCell.setCellStyle(borderStyle);

                rowIndex++;
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export DeliverySchedule data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    
}
