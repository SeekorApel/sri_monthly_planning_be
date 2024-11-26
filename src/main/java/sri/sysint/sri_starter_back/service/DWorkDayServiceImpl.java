package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
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

import sri.sysint.sri_starter_back.model.DWorkDay;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.repository.DWorkDayRepo;

@Service
@Transactional
public class DWorkDayServiceImpl {

    @Autowired
    private DWorkDayRepo dWorkDayRepo;

    public DWorkDayServiceImpl(DWorkDayRepo dWorkDayRepo) {
        this.dWorkDayRepo = dWorkDayRepo;
    }

    public BigDecimal getNewId() {
        return dWorkDayRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<DWorkDay> getAllDWorkDays() {
        Iterable<DWorkDay> workDays = dWorkDayRepo.findAll();
        List<DWorkDay> workDayList = new ArrayList<>();
        for (DWorkDay item : workDays) {
            workDayList.add(new DWorkDay(item));
        }
        return workDayList;
    }

    public Optional<DWorkDay> getDWorkDayById(BigDecimal id) {
        return dWorkDayRepo.findById(id);
    }
    
    public List<DWorkDay> getDWorkDayByDate(Date date) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String formattedDate = dateFormat.format(date);
        return dWorkDayRepo.findByDate(formattedDate);
    }


    public DWorkDay saveDWorkDay(DWorkDay workDay) {
        try {
            workDay.setDETAIL_WD_ID(getNewId());
            workDay.setSTATUS(BigDecimal.valueOf(1));
            workDay.setCREATION_DATE(new Date());
            workDay.setLAST_UPDATE_DATE(new Date());
            return dWorkDayRepo.save(workDay);
        } catch (Exception e) {
            System.err.println("Error saving work day: " + e.getMessage());
            throw e;
        }
    }

    public DWorkDay updateDWorkDay(DWorkDay workDay) {
        try {
            Optional<DWorkDay> currentWorkDayOpt = dWorkDayRepo.findById(workDay.getDETAIL_WD_ID());
            if (currentWorkDayOpt.isPresent()) {
                DWorkDay currentWorkDay = currentWorkDayOpt.get();
                currentWorkDay.setDATE_WD(workDay.getDATE_WD());
                currentWorkDay.setDESCRIPTION(workDay.getDESCRIPTION());
                currentWorkDay.setPARENT(workDay.getPARENT());
                currentWorkDay.setSTATUS(workDay.getSTATUS());
                currentWorkDay.setLAST_UPDATE_DATE(new Date());
                currentWorkDay.setLAST_UPDATED_BY(workDay.getLAST_UPDATED_BY());
                return dWorkDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("Work Day with ID " + workDay.getDETAIL_WD_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating work day: " + e.getMessage());
            throw e;
        }
    }

    public DWorkDay deleteDWorkDay(DWorkDay workDay) {
        try {
            Optional<DWorkDay> currentWorkDayOpt = dWorkDayRepo.findById(workDay.getDETAIL_WD_ID());
            if (currentWorkDayOpt.isPresent()) {
                DWorkDay currentWorkDay = currentWorkDayOpt.get();
                currentWorkDay.setSTATUS(BigDecimal.valueOf(0));
                currentWorkDay.setLAST_UPDATE_DATE(new Date());
                currentWorkDay.setLAST_UPDATED_BY(workDay.getLAST_UPDATED_BY());
                return dWorkDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("Work Day with ID " + workDay.getDETAIL_WD_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error deleting work day: " + e.getMessage());
            throw e;
        }
    }
    
    public DWorkDay restoreDWorkDay(DWorkDay workDay) {
        try {
            Optional<DWorkDay> currentWorkDayOpt = dWorkDayRepo.findById(workDay.getDETAIL_WD_ID());
            if (currentWorkDayOpt.isPresent()) {
                DWorkDay currentWorkDay = currentWorkDayOpt.get();
                currentWorkDay.setSTATUS(BigDecimal.valueOf(1));
                currentWorkDay.setLAST_UPDATE_DATE(new Date());
                currentWorkDay.setLAST_UPDATED_BY(workDay.getLAST_UPDATED_BY());
                return dWorkDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("Work Day with ID " + workDay.getDETAIL_WD_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error deleting work day: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllDWorkDay() {
    	dWorkDayRepo.deleteAll();
    }

    public ByteArrayInputStream exportDWorkDaysExcel() throws IOException {
        List<DWorkDay> workDays = dWorkDayRepo.findAll();
        return dataToExcel(workDays);
    }

    public ByteArrayInputStream dataToExcel(List<DWorkDay> workDays) throws IOException {
        String[] header = {
            "NOMOR",
            "DETAIL_WD_ID",
            "DATE_WD",
            "PARENT",
            "DESCRIPTION",
        };

        Workbook workbook = new XSSFWorkbook();  // Create a new workbook
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("WORK DAY DATA");

            // Create header font and style
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            // Create border style for both header and data cells
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            // Create header style
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Set dynamic column width
            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 15 * 256); // Set a reasonable default width for all columns
            }

            // Create the header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            // Fill data rows
            int rowIndex = 1;
            int nomor = 1;
            SimpleDateFormat dateFormat = new SimpleDateFormat("E MMMM dd, yyyy");  // Define date format

            for (DWorkDay workDay : workDays) {
                Row dataRow = sheet.createRow(rowIndex++);

                // Create cells and set values
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(workDay.getDETAIL_WD_ID().doubleValue());  // Assuming DETAIL_WD_ID is BigDecimal
                idCell.setCellStyle(borderStyle);

                Cell dateCell = dataRow.createCell(2);
                if (workDay.getDATE_WD() != null) {
                    String formattedDate = dateFormat.format(workDay.getDATE_WD());  // Format date
                    dateCell.setCellValue(formattedDate);
                } else {
                    dateCell.setCellValue("");  // If DATE_WD is null, leave it empty
                }
                dateCell.setCellStyle(borderStyle);

                Cell parentCell = dataRow.createCell(3);
                parentCell.setCellValue(workDay.getPARENT());
                parentCell.setCellStyle(borderStyle);

                Cell descriptionCell = dataRow.createCell(4);
                descriptionCell.setCellValue(workDay.getDESCRIPTION());
                descriptionCell.setCellStyle(borderStyle);
            }

            // Write the data to the output stream
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());  // Return the ByteArrayInputStream for download
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export data");
            return null;
        } finally {
            try {
                workbook.close();  // Ensure the workbook is closed
                out.close();  // Ensure the output stream is closed
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    
}
