package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.TimeZone;
import java.util.stream.Stream;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.DWorkDayHoursSpesific;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.repository.DWorkDayHoursSpecificRepo;
import sri.sysint.sri_starter_back.repository.WorkDayRepo;

@Service
@Transactional
public class DWorkDayHoursSpecificServiceImpl {

	 @Autowired
	 private DWorkDayHoursSpecificRepo dWorkDayHoursSpecificRepo;
	 
	 @Autowired
	 private WorkDayRepo workDayRepo;

	    public DWorkDayHoursSpecificServiceImpl(DWorkDayHoursSpecificRepo dWorkDayHoursRepo) {
	        this.dWorkDayHoursSpecificRepo = dWorkDayHoursRepo;
	    }

	    // Mengambil ID baru untuk entri
	    public BigDecimal getNewSpecificId() {
	        return dWorkDayHoursSpecificRepo.getNewId().add(BigDecimal.ONE);
	    }

	    // Mengambil semua entri jam kerja spesifik
	    @Transactional
	    public List<DWorkDayHoursSpesific> getAllWorkDayHoursSpecific() {
	        return dWorkDayHoursSpecificRepo.findAll();
	    }

	    // Mengambil jam kerja berdasarkan tanggal
	    public Optional<DWorkDayHoursSpesific> getWorkDayHoursSpecificByDate(Date date) {
	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
	        String formattedDate = dateFormat.format(date);

	        System.out.println("Formatted date being sent to repository: " + formattedDate);
	        return dWorkDayHoursSpecificRepo.findDWdHoursByDate(formattedDate);
	    }

	    // Mengambil jam kerja berdasarkan tanggal dan deskripsi
	    public Optional<DWorkDayHoursSpesific> getWorkDayHoursSpecificByDateDesc(Date date, String description) {
	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
	        String formattedDate = dateFormat.format(date);

	        System.out.println("Formatted date and description being sent to repository: " + formattedDate + ", " + description);
	        return dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(formattedDate, description);
	    }
	    
	    // Fungsi untuk mengambil jam kerja berdasarkan bulan dan tahun
	    public List<DWorkDayHoursSpesific> getWorkDayHoursByMonthAndYear(int month, int year) {
	        System.out.println("Getting work day hours for month: " + month + " and year: " + year);
	        return dWorkDayHoursSpecificRepo.findDWdHoursByMonthAndYear(month, year);
	    }

	    public DWorkDayHoursSpesific saveWorkDayHoursSpecific(DWorkDayHoursSpesific workHoursSpecific) {
	        try {
	            // Menghitung total waktu per shift
	            workHoursSpecific.setSHIFT1_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT1_START_TIME(), workHoursSpecific.getSHIFT1_END_TIME()));
	            workHoursSpecific.setSHIFT2_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT2_START_TIME(), workHoursSpecific.getSHIFT2_END_TIME()));
	            workHoursSpecific.setSHIFT3_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT3_START_TIME(), workHoursSpecific.getSHIFT3_END_TIME()));

	            workHoursSpecific.setDETAIL_WD_HOURS_SPECIFIC_ID(getNewSpecificId());
	            workHoursSpecific.setCREATION_DATE(new Date());
	            workHoursSpecific.setLAST_UPDATE_DATE(new Date());
	            workHoursSpecific.setSTATUS(BigDecimal.ONE);

	            return dWorkDayHoursSpecificRepo.save(workHoursSpecific);
	        } catch (Exception e) {
	            System.err.println("Error saving work hours: " + e.getMessage());
	            throw e;
	        }
	    }


	    public DWorkDayHoursSpesific updateWorkDayHoursSpecific(DWorkDayHoursSpesific workHoursSpecific) {
	        try {
	            Optional<DWorkDayHoursSpesific> currentWorkHoursSpecificOpt =
	                    dWorkDayHoursSpecificRepo.findById(workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID());

	            if (currentWorkHoursSpecificOpt.isPresent()) {
	                DWorkDayHoursSpesific currentWorkHoursSpecific = currentWorkHoursSpecificOpt.get();

	                currentWorkHoursSpecific.setSHIFT1_START_TIME(workHoursSpecific.getSHIFT1_START_TIME());
	                currentWorkHoursSpecific.setSHIFT1_END_TIME(workHoursSpecific.getSHIFT1_END_TIME());
	                currentWorkHoursSpecific.setSHIFT2_START_TIME(workHoursSpecific.getSHIFT2_START_TIME());
	                currentWorkHoursSpecific.setSHIFT2_END_TIME(workHoursSpecific.getSHIFT2_END_TIME());
	                currentWorkHoursSpecific.setSHIFT3_START_TIME(workHoursSpecific.getSHIFT3_START_TIME());
	                currentWorkHoursSpecific.setSHIFT3_END_TIME(workHoursSpecific.getSHIFT3_END_TIME());
	                // Menghitung total waktu per shift
	                currentWorkHoursSpecific.setSHIFT1_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT1_START_TIME(), workHoursSpecific.getSHIFT1_END_TIME()));
	                currentWorkHoursSpecific.setSHIFT2_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT2_START_TIME(), workHoursSpecific.getSHIFT2_END_TIME()));
	                currentWorkHoursSpecific.setSHIFT3_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT3_START_TIME(), workHoursSpecific.getSHIFT3_END_TIME()));

	                currentWorkHoursSpecific.setLAST_UPDATE_DATE(new Date());
	                currentWorkHoursSpecific.setLAST_UPDATED_BY(workHoursSpecific.getLAST_UPDATED_BY());

	                return dWorkDayHoursSpecificRepo.save(currentWorkHoursSpecific);
	            } else {
	                throw new RuntimeException("WorkHoursSpecific with ID " + workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID() + " not found.");
	            }
	        } catch (Exception e) {
	            System.err.println("Error updating work hours: " + e.getMessage());
	            throw e;
	        }
	    }


	    // Menghapus entri jam kerja spesifik
	    public DWorkDayHoursSpesific deleteWorkDayHoursSpecific(DWorkDayHoursSpesific workHoursSpecific) {
	        try {
	            Optional<DWorkDayHoursSpesific> currentWorkHoursSpecificOpt =
	                    dWorkDayHoursSpecificRepo.findById(workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID());

	            if (currentWorkHoursSpecificOpt.isPresent()) {
	                DWorkDayHoursSpesific currentWorkHoursSpecific = currentWorkHoursSpecificOpt.get();
	                currentWorkHoursSpecific.setSTATUS(BigDecimal.ZERO); // Menandakan tidak aktif
	                currentWorkHoursSpecific.setLAST_UPDATE_DATE(new Date());
	                currentWorkHoursSpecific.setLAST_UPDATED_BY(workHoursSpecific.getLAST_UPDATED_BY());

	                return dWorkDayHoursSpecificRepo.save(currentWorkHoursSpecific);
	            } else {
	                throw new RuntimeException("WorkHoursSpecific with ID " + workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID() + " not found.");
	            }
	        } catch (Exception e) {
	            System.err.println("Error deleting work hours: " + e.getMessage());
	            throw e;
	        }
	    }

	    // Mengembalikan (restore) entri jam kerja spesifik
	    public DWorkDayHoursSpesific restoreWorkDayHoursSpecific(DWorkDayHoursSpesific workHoursSpecific) {
	        try {
	            Optional<DWorkDayHoursSpesific> currentWorkHoursSpecificOpt =
	                    dWorkDayHoursSpecificRepo.findById(workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID());

	            if (currentWorkHoursSpecificOpt.isPresent()) {
	                DWorkDayHoursSpesific currentWorkHoursSpecific = currentWorkHoursSpecificOpt.get();
	                currentWorkHoursSpecific.setSTATUS(BigDecimal.ONE); // Mengaktifkan kembali
	                currentWorkHoursSpecific.setLAST_UPDATE_DATE(new Date());
	                currentWorkHoursSpecific.setLAST_UPDATED_BY(workHoursSpecific.getLAST_UPDATED_BY());

	                return dWorkDayHoursSpecificRepo.save(currentWorkHoursSpecific);
	            } else {
	                throw new RuntimeException("WorkHoursSpecific with ID " + workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID() + " not found.");
	            }
	        } catch (Exception e) {
	            System.err.println("Error restoring work hours: " + e.getMessage());
	            throw e;
	        }
	    }

	    // Menghapus semua entri jam kerja spesifik
	    public void deleteAllWorkDayHoursSpecific() {
	        dWorkDayHoursSpecificRepo.deleteAll();
	    }

	 // Mengganti implementasi untuk menggabungkan waktu tanpa melibatkan Date
	    private int combineTimeWithMinutes(String timeStr) {
	        // Mengonversi string waktu "HH:mm" menjadi total menit
	        String[] parts = timeStr.split(":");
	        int hours = Integer.parseInt(parts[0]);
	        int minutes = Integer.parseInt(parts[1]);
	        return hours * 60 + minutes;
	    }
	 // Menghitung total waktu dalam menit
	    private BigDecimal calculateTotalTime(String startTimeStr, String endTimeStr) {
	        int startMinutes = combineTimeWithMinutes(startTimeStr);
	        int endMinutes = combineTimeWithMinutes(endTimeStr);
	        
	        // Menghitung selisih waktu dalam menit
	        int diffMinutes = endMinutes - startMinutes;
	        
	        // Jika waktu berakhir sebelum waktu mulai (misalnya lewat tengah malam), tambahkan 1440 menit (24 jam)
	        if (diffMinutes < 0) {
	            diffMinutes += 24 * 60; // Tambahkan 24 jam dalam menit
	        }

	        // Kembalikan hasil dalam menit
	        return BigDecimal.valueOf(diffMinutes); // Menghitung dalam menit
	    }



    // Export to Excel function
    public ByteArrayInputStream exportWorkDayHoursSpecificExcel() throws IOException {
        List<DWorkDayHoursSpesific> workHoursSpecificList = dWorkDayHoursSpecificRepo.findAll();
        return dataToExcel(workHoursSpecificList);
    }

    private ByteArrayInputStream dataToExcel(List<DWorkDayHoursSpesific> workHoursSpecificList) throws IOException {
        List<String> columns = List.of(
            "No.", "DETAIL_WD_HOURS_SPECIFIC_ID", "DATE_WD", 
            "SHIFT1_START_TIME", "SHIFT1_END_TIME", "SHIFT1_TOTAL_TIME",
            "SHIFT2_START_TIME", "SHIFT2_END_TIME", "SHIFT2_TOTAL_TIME",
            "SHIFT3_START_TIME", "SHIFT3_END_TIME", "SHIFT3_TOTAL_TIME",
            "DESCRIPTION"
        );

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("WorkDayHoursSpecific");

            // Font and style for header
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            // Create border style
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            // Header style with yellow background and border
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Fill header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns.get(i));
                cell.setCellStyle(headerStyle);
                sheet.setColumnWidth(i, 15 * 256);
            }

            // Fill data rows
            SimpleDateFormat dateFormat = new SimpleDateFormat("E MMMM dd, yyyy");
            SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
            for (int i = 0; i < workHoursSpecificList.size(); i++) {
            	DWorkDayHoursSpesific workHoursSpecific = workHoursSpecificList.get(i);
                Row dataRow = sheet.createRow(i + 1);

                // "No." column
                Cell noCell = dataRow.createCell(0);
                noCell.setCellValue(i + 1); 
                noCell.setCellStyle(borderStyle);

                // DETAIL_WD_HOURS_SPECIFIC_ID
                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                // DATE_WD
                Cell dateCell = dataRow.createCell(2);
                dateCell.setCellValue(dateFormat.format(workHoursSpecific.getDATE_WD()));
                dateCell.setCellStyle(borderStyle);

                // SHIFT1 details
                dataRow.createCell(3).setCellValue(workHoursSpecific.getSHIFT1_START_TIME());
                dataRow.createCell(4).setCellValue(workHoursSpecific.getSHIFT1_END_TIME());
                dataRow.createCell(5).setCellValue(workHoursSpecific.getSHIFT1_TOTAL_TIME() != null ? workHoursSpecific.getSHIFT1_TOTAL_TIME().doubleValue() : 0);

                // SHIFT2 details
                dataRow.createCell(6).setCellValue(workHoursSpecific.getSHIFT2_START_TIME());
                dataRow.createCell(7).setCellValue(workHoursSpecific.getSHIFT2_END_TIME());
                dataRow.createCell(8).setCellValue(workHoursSpecific.getSHIFT2_TOTAL_TIME() != null ? workHoursSpecific.getSHIFT2_TOTAL_TIME().doubleValue() : 0);

                // SHIFT3 details
                dataRow.createCell(9).setCellValue(workHoursSpecific.getSHIFT3_START_TIME());
                dataRow.createCell(10).setCellValue(workHoursSpecific.getSHIFT3_END_TIME());
                dataRow.createCell(11).setCellValue(workHoursSpecific.getSHIFT3_TOTAL_TIME() != null ? workHoursSpecific.getSHIFT3_TOTAL_TIME().doubleValue() : 0);

                // DESCRIPTION
                dataRow.createCell(12).setCellValue(workHoursSpecific.getDESCRIPTION() != null ? workHoursSpecific.getDESCRIPTION() : "");

                // Apply border style for each cell in the row
                for (int j = 0; j < columns.size(); j++) {
                    dataRow.getCell(j).setCellStyle(borderStyle);
                }
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } finally {
            workbook.close();
            out.close();
        }
    }

    public Optional<DWorkDayHoursSpesific> updateShiftTimes(String startTime, String endTime, String parsedDate, String description, int shift) {
        Optional<DWorkDayHoursSpesific> workHoursSpecificOpt = dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(parsedDate, description);

        if (workHoursSpecificOpt.isPresent()) {
            DWorkDayHoursSpesific workHoursSpecific = workHoursSpecificOpt.get();

            switch (shift) {
                case 1:
                    workHoursSpecific.setSHIFT1_START_TIME(startTime);
                    workHoursSpecific.setSHIFT1_END_TIME(endTime);
                    workHoursSpecific.setSHIFT1_TOTAL_TIME(calculateTotalTime(startTime, endTime));
                    break;
                case 2:
                    workHoursSpecific.setSHIFT2_START_TIME(startTime);
                    workHoursSpecific.setSHIFT2_END_TIME(endTime);
                    workHoursSpecific.setSHIFT2_TOTAL_TIME(calculateTotalTime(startTime, endTime));
                    break;
                case 3:
                    workHoursSpecific.setSHIFT3_START_TIME(startTime);
                    workHoursSpecific.setSHIFT3_END_TIME(endTime);
                    workHoursSpecific.setSHIFT3_TOTAL_TIME(calculateTotalTime(startTime, endTime));
                    break;
                default:
                    throw new IllegalArgumentException("Invalid shift specified: " + shift);
            }

            workHoursSpecific.setLAST_UPDATE_DATE(new Date());
            dWorkDayHoursSpecificRepo.save(workHoursSpecific);
            
            updateOffAndSemiOff(workHoursSpecific, parsedDate, description);

            System.out.println("Shift " + shift + " updated successfully.");
            return Optional.of(workHoursSpecific);
        } else {
            System.out.println("No WorkDayHoursSpecific record found for date: " + parsedDate);
            return Optional.empty();
        }
    }
   

    
    private void updateOffAndSemiOff(DWorkDayHoursSpesific dWorkDayHoursSpecific, String dateWd, String description) {
        Optional<WorkDay> workDayOpt = workDayRepo.findByDDateWd(dateWd);

        workDayOpt.ifPresentOrElse(workDay -> {
            // Ambil total waktu dari setiap shift
            BigDecimal shift1 = dWorkDayHoursSpecific.getSHIFT1_TOTAL_TIME() != null ? dWorkDayHoursSpecific.getSHIFT1_TOTAL_TIME() : BigDecimal.ZERO;
            BigDecimal shift2 = dWorkDayHoursSpecific.getSHIFT2_TOTAL_TIME() != null ? dWorkDayHoursSpecific.getSHIFT2_TOTAL_TIME() : BigDecimal.ZERO;
            BigDecimal shift3 = dWorkDayHoursSpecific.getSHIFT3_TOTAL_TIME() != null ? dWorkDayHoursSpecific.getSHIFT3_TOTAL_TIME() : BigDecimal.ZERO;

            // Log nilai shift
            System.out.println("Shift1: " + shift1);
            System.out.println("Shift2: " + shift2);
            System.out.println("Shift3: " + shift3);

            DayOfWeek dayOfWeek = getDayOfWeekFromDate(dateWd);

            // Apply logic based on the description parameter
            if ("WD_NORMAL".equals(description)) {
                // Set IWD_SHIFT_1, IWD_SHIFT_2, IWD_SHIFT_3 based on shift times
                workDay.setIWD_SHIFT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                workDay.setIWD_SHIFT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                workDay.setIWD_SHIFT_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                System.out.println("IWD_SHIFT_1: " + workDay.getIWD_SHIFT_1());
                System.out.println("IWD_SHIFT_2: " + workDay.getIWD_SHIFT_2());
                System.out.println("IWD_SHIFT_3: " + workDay.getIWD_SHIFT_3());
            } else if ("OT_TT".equals(description)) {
                // Apply logic for OT_TT (Overtime Total Time)
                workDay.setIOT_TT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                workDay.setIOT_TT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                workDay.setIOT_TT_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                System.out.println("IOT_TT_1: " + workDay.getIOT_TT_1());
                System.out.println("IOT_TT_2: " + workDay.getIOT_TT_2());
                System.out.println("IOT_TT_3: " + workDay.getIOT_TT_3());
            } else if ("OT_TL".equals(description)) {
                // Apply logic for OT_TL (Overtime Total Leave)
                workDay.setIOT_TL_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                workDay.setIOT_TL_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                workDay.setIOT_TL_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                System.out.println("IOT_TL_1: " + workDay.getIOT_TL_1());
                System.out.println("IOT_TL_2: " + workDay.getIOT_TL_2());
                System.out.println("IOT_TL_3: " + workDay.getIOT_TL_3());
            }

            // Log and update OFF and SEMI_OFF based on the day of the week and shift times
            if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else if (dayOfWeek == DayOfWeek.MONDAY) {
                if (shift1.compareTo(BigDecimal.ZERO) > 0 && shift2.compareTo(BigDecimal.ZERO) > 0) {
                    workDay.setOFF(BigDecimal.ZERO);
                    workDay.setSEMI_OFF(BigDecimal.ZERO);
                } else if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0) {
                    workDay.setOFF(BigDecimal.ONE);
                    workDay.setSEMI_OFF(BigDecimal.ZERO);
                } else {
                    workDay.setOFF(BigDecimal.ZERO);
                    workDay.setSEMI_OFF(BigDecimal.ONE);
                }
            } else if (dayOfWeek.getValue() >= DayOfWeek.TUESDAY.getValue() && dayOfWeek.getValue() <= DayOfWeek.FRIDAY.getValue()) {
                if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0 && shift3.compareTo(BigDecimal.ZERO) == 0) {
                    workDay.setOFF(BigDecimal.ONE);
                    workDay.setSEMI_OFF(BigDecimal.ZERO);
                } else if (shift1.compareTo(BigDecimal.ZERO) > 0 && shift2.compareTo(BigDecimal.ZERO) > 0 && shift3.compareTo(BigDecimal.ZERO) > 0) {
                    workDay.setOFF(BigDecimal.ZERO);
                    workDay.setSEMI_OFF(BigDecimal.ZERO);
                } else {
                    workDay.setOFF(BigDecimal.ZERO);
                    workDay.setSEMI_OFF(BigDecimal.ONE);
                }
            }

            System.out.println("SEMI_OFF value: " + workDay.getSEMI_OFF());
            System.out.println("OFF value: " + workDay.getOFF());

            workDayRepo.save(workDay);
        }, () -> {
            throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
        });
    }


    private DayOfWeek getDayOfWeekFromDate(String dateWd) {
        return LocalDate.parse(dateWd, DateTimeFormatter.ofPattern("dd-MM-yyyy")).getDayOfWeek();
    }

}
