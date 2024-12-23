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
import sri.sysint.sri_starter_back.model.Setting;
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

	private DWorkDayHoursSpesific currentWorkHoursSpecific;

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
	    
	    // Mengambil jam kerja berdasarkan tanggal list
	    public List<DWorkDayHoursSpesific> getWorkDayHoursSpecificListByDate(Date date) {
	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
	        String formattedDate = dateFormat.format(date);

	        System.out.println("Formatted date being sent to repository: " + formattedDate);
	        return dWorkDayHoursSpecificRepo.findDWdHoursListByDate(formattedDate);
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

	    public DWorkDayHoursSpesific saveWorkDayHoursSpecific(DWorkDayHoursSpesific workHoursSpecific) throws Exception {
	        try {
	        	
                BigDecimal shift1TotalTime = calculateTotalTime(workHoursSpecific.getSHIFT1_START_TIME(), workHoursSpecific.getSHIFT1_END_TIME());

                Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                calendar.setTime(workHoursSpecific.getDATE_WD());

                System.out.println("Day of Week: " + calendar.get(Calendar.DAY_OF_WEEK)); 

                if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.FRIDAY) {
                    Calendar baseDate = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta")); // Set zona waktu ke UTC
                    baseDate.setTime(workHoursSpecific.getDATE_WD()); // Menggunakan tanggal dari workHoursSpecific
                    Calendar startTime = convertToCalendar(workHoursSpecific.getSHIFT1_START_TIME(), baseDate);
                    Calendar endTime = convertToCalendar(workHoursSpecific.getSHIFT1_END_TIME(), baseDate);

                    Calendar breakStart = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                    breakStart.setTime(workHoursSpecific.getDATE_WD());
                    breakStart.set(Calendar.HOUR_OF_DAY, 11); // Jam 11
                    breakStart.set(Calendar.MINUTE, 40);     // Menit 40
                    breakStart.set(Calendar.SECOND, 0);      // Detik 0

                    Calendar breakEnd = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                    breakEnd.setTime(workHoursSpecific.getDATE_WD());
                    breakEnd.set(Calendar.HOUR_OF_DAY, 12);   // Jam 12
                    breakEnd.set(Calendar.MINUTE, 40);        // Menit 40
                    breakEnd.set(Calendar.SECOND, 0);         // Detik 0

                    // Konversi ke zona waktu WIB jika diperlukan
                    TimeZone wibTimeZone = TimeZone.getTimeZone("Asia/Jakarta");
                    breakStart.setTimeZone(wibTimeZone);
                    breakEnd.setTimeZone(wibTimeZone);

                    // Debug log untuk memastikan jam sudah diatur dengan benar
                    System.out.println("Break Start: " + breakStart.getTime()); // Debugging
                    System.out.println("Break End: " + breakEnd.getTime());   // Debugging


                    System.out.println("start time: " + startTime.getTime()); // Debug log
                    System.out.println("end time: " + endTime.getTime()); // Debug log

                    // Jika waktu mulai dan selesai melewati jam istirahat
                    if (startTime.before(breakEnd) && endTime.after(breakStart)) {
                        long overlapStart = Math.max(startTime.getTimeInMillis(), breakStart.getTimeInMillis());
                        long overlapEnd = Math.min(endTime.getTimeInMillis(), breakEnd.getTimeInMillis());
                        if (overlapStart < overlapEnd) {
                            long overlapDurationMillis = overlapEnd - overlapStart;
                            long overlapDurationMinutes = overlapDurationMillis / 60000;
                            shift1TotalTime = shift1TotalTime.subtract(new BigDecimal(overlapDurationMinutes));
                            System.out.println("Updated shift1TotalTime: " + shift1TotalTime); // Debugging
                        }
                    }
                }

                workHoursSpecific.setSHIFT1_TOTAL_TIME(shift1TotalTime);

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
	    
//	    private Calendar convertToCalendar(String timeString, Calendar baseDate) throws Exception {
//	        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
//	        Date date = sdf.parse(timeString); // Mengonversi string ke Date
//
//	        // Gunakan tanggal dari baseDate (tanggal kerja) untuk menetapkan waktu yang tepat
//	        Calendar calendar = Calendar.getInstance();
//	        calendar.setTime(date); // Set waktu dari date yang sudah di-parsing
//
//	        // Set tanggal dari baseDate ke calendar
//	        calendar.set(Calendar.YEAR, baseDate.get(Calendar.YEAR));
//	        calendar.set(Calendar.MONTH, baseDate.get(Calendar.MONTH));
//	        calendar.set(Calendar.DAY_OF_MONTH, baseDate.get(Calendar.DAY_OF_MONTH));
//
//	        // Debugging untuk memastikan waktu yang dikonversi benar
//	        System.out.println("Converted time for " + timeString + ": " + calendar.getTime()); 
//	        return calendar;
//	    }
	    
	    public String formatDateToString(Date date) {
	        SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
	        return formatter.format(date);
	    }
	    
	    private Calendar convertToCalendar(String timeString, Calendar baseDate) throws Exception {
	        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
	        Date date = sdf.parse(timeString); // Mengonversi string ke Date
	        
	        // Gunakan tanggal dari baseDate (tanggal kerja) untuk menetapkan waktu yang tepat
	        Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("UTC")); // Set Zona Waktu ke UTC
	        calendar.setTime(date); // Set waktu dari date yang sudah di-parsing

	        // Set tanggal dari baseDate ke calendar
	        calendar.set(Calendar.YEAR, baseDate.get(Calendar.YEAR));
	        calendar.set(Calendar.MONTH, baseDate.get(Calendar.MONTH));
	        calendar.set(Calendar.DAY_OF_MONTH, baseDate.get(Calendar.DAY_OF_MONTH));

	        // Debugging untuk memastikan waktu yang dikonversi benar
	        System.out.println("Converted time for " + timeString + ": " + calendar.getTime()); 
	        return calendar;
	    }

	    public DWorkDayHoursSpesific updateWorkDayHoursSpecific(DWorkDayHoursSpesific workHoursSpecific) throws Exception {
	        try {
	            String formattedDate = formatDateToString(workHoursSpecific.getDATE_WD());

	            Optional<DWorkDayHoursSpesific> currentWorkHoursSpecificOpt =
	                    dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(formattedDate, workHoursSpecific.getDESCRIPTION());

	            if (currentWorkHoursSpecificOpt.isPresent()) {
	                DWorkDayHoursSpesific currentWorkHoursSpecific = currentWorkHoursSpecificOpt.get();

	                // Update shift times
	                currentWorkHoursSpecific.setSHIFT1_START_TIME(workHoursSpecific.getSHIFT1_START_TIME());
	                currentWorkHoursSpecific.setSHIFT1_END_TIME(workHoursSpecific.getSHIFT1_END_TIME());
	                currentWorkHoursSpecific.setSHIFT2_START_TIME(workHoursSpecific.getSHIFT2_START_TIME());
	                currentWorkHoursSpecific.setSHIFT2_END_TIME(workHoursSpecific.getSHIFT2_END_TIME());
	                currentWorkHoursSpecific.setSHIFT3_START_TIME(workHoursSpecific.getSHIFT3_START_TIME());
	                currentWorkHoursSpecific.setSHIFT3_END_TIME(workHoursSpecific.getSHIFT3_END_TIME());

	                // Menghitung total waktu per shift
	                BigDecimal shift1TotalTime = calculateTotalTime(workHoursSpecific.getSHIFT1_START_TIME(), workHoursSpecific.getSHIFT1_END_TIME());

	                // Menggunakan zona waktu UTC
	                Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
	                calendar.setTime(workHoursSpecific.getDATE_WD());

	                System.out.println("Day of Week: " + calendar.get(Calendar.DAY_OF_WEEK)); // Debug log untuk memastikan hari yang benar

	                if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.FRIDAY) {
	                    // Mengonversi start dan end time dari String ke Calendar menggunakan baseDate
	                    Calendar baseDate = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta")); // Set zona waktu ke UTC
	                    baseDate.setTime(workHoursSpecific.getDATE_WD()); // Menggunakan tanggal dari workHoursSpecific
	                    Calendar startTime = convertToCalendar(workHoursSpecific.getSHIFT1_START_TIME(), baseDate);
	                    Calendar endTime = convertToCalendar(workHoursSpecific.getSHIFT1_END_TIME(), baseDate);

	                 // Menggunakan UTC untuk waktu, tetapi setelah itu konversi kembali ke WIB jika perlu
	                    Calendar breakStart = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
	                    breakStart.setTime(workHoursSpecific.getDATE_WD());
	                    breakStart.set(Calendar.HOUR_OF_DAY, 11); // Jam 11
	                    breakStart.set(Calendar.MINUTE, 40);     // Menit 40
	                    breakStart.set(Calendar.SECOND, 0);      // Detik 0

	                    Calendar breakEnd = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
	                    breakEnd.setTime(workHoursSpecific.getDATE_WD());
	                    breakEnd.set(Calendar.HOUR_OF_DAY, 12);   // Jam 12
	                    breakEnd.set(Calendar.MINUTE, 40);        // Menit 40
	                    breakEnd.set(Calendar.SECOND, 0);         // Detik 0

	                    // Konversi ke zona waktu WIB jika diperlukan
	                    TimeZone wibTimeZone = TimeZone.getTimeZone("Asia/Jakarta");
	                    breakStart.setTimeZone(wibTimeZone);
	                    breakEnd.setTimeZone(wibTimeZone);

	                    // Debug log untuk memastikan jam sudah diatur dengan benar
	                    System.out.println("Break Start: " + breakStart.getTime()); // Debugging
	                    System.out.println("Break End: " + breakEnd.getTime());   // Debugging


	                    System.out.println("start time: " + startTime.getTime()); // Debug log
	                    System.out.println("end time: " + endTime.getTime()); // Debug log

	                    // Jika waktu mulai dan selesai melewati jam istirahat
	                    if (startTime.before(breakEnd) && endTime.after(breakStart)) {
	                        long overlapStart = Math.max(startTime.getTimeInMillis(), breakStart.getTimeInMillis());
	                        long overlapEnd = Math.min(endTime.getTimeInMillis(), breakEnd.getTimeInMillis());
	                        if (overlapStart < overlapEnd) {
	                            long overlapDurationMillis = overlapEnd - overlapStart;
	                            long overlapDurationMinutes = overlapDurationMillis / 60000;
	                            shift1TotalTime = shift1TotalTime.subtract(new BigDecimal(overlapDurationMinutes));
	                            System.out.println("Updated shift1TotalTime: " + shift1TotalTime); // Debugging
	                        }
	                    }
	                }

	                currentWorkHoursSpecific.setSHIFT1_TOTAL_TIME(shift1TotalTime);
	                currentWorkHoursSpecific.setSHIFT2_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT2_START_TIME(), workHoursSpecific.getSHIFT2_END_TIME()));
	                currentWorkHoursSpecific.setSHIFT3_TOTAL_TIME(calculateTotalTime(workHoursSpecific.getSHIFT3_START_TIME(), workHoursSpecific.getSHIFT3_END_TIME()));

	                // Set last update date and user
	                currentWorkHoursSpecific.setLAST_UPDATE_DATE(new Date());
	                currentWorkHoursSpecific.setLAST_UPDATED_BY(workHoursSpecific.getLAST_UPDATED_BY());

	                System.out.println("Masuk sini 0");

	                // Save and return the updated record
	                DWorkDayHoursSpesific updatedWorkHoursSpecific = dWorkDayHoursSpecificRepo.save(currentWorkHoursSpecific);

	                System.out.println("Masuk sini 1");
	                
	                // Call updateOffAndSemiOff method
	                updateOffAndSemiOff(updatedWorkHoursSpecific, formattedDate, workHoursSpecific.getDESCRIPTION());

	                System.out.println("Masuk sini 2");

	                return updatedWorkHoursSpecific;
	            } else {
	                throw new RuntimeException("WorkHoursSpecific with ID " + workHoursSpecific.getDETAIL_WD_HOURS_SPECIFIC_ID() + " not found.");
	            }
	        } catch (Exception e) {
	            System.err.println("Error updating work hours: " + e.getMessage());
	            throw e;
	        }
	    }

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

    public Optional<DWorkDayHoursSpesific> updateShiftTimes(String startTime, String endTime, String parsedDate, String description, int shift) throws Exception {
        Optional<DWorkDayHoursSpesific> workHoursSpecificOpt = dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(parsedDate, description);

        if (workHoursSpecificOpt.isPresent()) {
            DWorkDayHoursSpesific workHoursSpecific = workHoursSpecificOpt.get();

            switch (shift) {
                case 1:
                	BigDecimal shift1TotalTime = calculateTotalTime(startTime, endTime);

                    Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                    calendar.setTime(workHoursSpecific.getDATE_WD());

                    System.out.println("Day of Week: " + calendar.get(Calendar.DAY_OF_WEEK)); 

                    if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.FRIDAY) {
                        // Mengonversi start dan end time dari String ke Calendar menggunakan baseDate
                        Calendar baseDate = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta")); 
                        baseDate.setTime(workHoursSpecific.getDATE_WD()); 
                        Calendar startTimee = convertToCalendar(workHoursSpecific.getSHIFT1_START_TIME(), baseDate);
                        Calendar endTimee = convertToCalendar(workHoursSpecific.getSHIFT1_END_TIME(), baseDate);

                        // Menggunakan UTC untuk waktu, tetapi setelah itu konversi kembali ke WIB jika perlu
                        Calendar breakStart = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                        breakStart.setTime(workHoursSpecific.getDATE_WD());
                        breakStart.set(Calendar.HOUR_OF_DAY, 11); // Jam 11
                        breakStart.set(Calendar.MINUTE, 40);     // Menit 40
                        breakStart.set(Calendar.SECOND, 0);      // Detik 0

                        Calendar breakEnd = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                        breakEnd.setTime(workHoursSpecific.getDATE_WD());
                        breakEnd.set(Calendar.HOUR_OF_DAY, 12);   // Jam 12
                        breakEnd.set(Calendar.MINUTE, 40);        // Menit 40
                        breakEnd.set(Calendar.SECOND, 0);         // Detik 0

                        // Konversi ke zona waktu WIB jika diperlukan
                        TimeZone wibTimeZone = TimeZone.getTimeZone("Asia/Jakarta");
                        breakStart.setTimeZone(wibTimeZone);
                        breakEnd.setTimeZone(wibTimeZone);

                        // Debug log untuk memastikan jam sudah diatur dengan benar
                        System.out.println("Break Start: " + breakStart.getTime()); 
                        System.out.println("Break End: " + breakEnd.getTime());   

                        System.out.println("start time: " + startTimee.getTime()); 
                        System.out.println("end time: " + endTimee.getTime()); 

                        // Jika waktu mulai dan selesai melewati jam istirahat
                        if (startTimee.before(breakEnd) && endTimee.after(breakStart)) {
                            long overlapStart = Math.max(startTimee.getTimeInMillis(), breakStart.getTimeInMillis());
                            long overlapEnd = Math.min(endTimee.getTimeInMillis(), breakEnd.getTimeInMillis());
                            if (overlapStart < overlapEnd) {
                                long overlapDurationMillis = overlapEnd - overlapStart;
                                long overlapDurationMinutes = overlapDurationMillis / 60000;
                                shift1TotalTime = shift1TotalTime.subtract(new BigDecimal(overlapDurationMinutes));
                                System.out.println("Updated shift1TotalTime: " + shift1TotalTime); 
                            }
                        }
                    }
		            
                    workHoursSpecific.setSHIFT1_START_TIME(startTime);
                    workHoursSpecific.setSHIFT1_END_TIME(endTime);
                    workHoursSpecific.setSHIFT1_TOTAL_TIME(shift1TotalTime);
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
    
    public Optional<DWorkDayHoursSpesific> updateShiftTimesforWd(String startTime, String endTime, String parsedDate, String description, int shift) throws Exception {
        Optional<DWorkDayHoursSpesific> workHoursSpecificOpt = dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(parsedDate, description);

        if (workHoursSpecificOpt.isPresent()) {
            DWorkDayHoursSpesific workHoursSpecific = workHoursSpecificOpt.get();

            switch (shift) {
                case 1:
                	BigDecimal shift1TotalTime = calculateTotalTime(startTime, endTime);

                    Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                    calendar.setTime(workHoursSpecific.getDATE_WD());

                    System.out.println("Day of Week: " + calendar.get(Calendar.DAY_OF_WEEK)); 

                    if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.FRIDAY) {
                        // Mengonversi start dan end time dari String ke Calendar menggunakan baseDate
                        Calendar baseDate = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta")); 
                        baseDate.setTime(workHoursSpecific.getDATE_WD()); 
                        Calendar startTimee = convertToCalendar(workHoursSpecific.getSHIFT1_START_TIME(), baseDate);
                        Calendar endTimee = convertToCalendar(workHoursSpecific.getSHIFT1_END_TIME(), baseDate);

                        // Menggunakan UTC untuk waktu, tetapi setelah itu konversi kembali ke WIB jika perlu
                        Calendar breakStart = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                        breakStart.setTime(workHoursSpecific.getDATE_WD());
                        breakStart.set(Calendar.HOUR_OF_DAY, 11); // Jam 11
                        breakStart.set(Calendar.MINUTE, 40);     // Menit 40
                        breakStart.set(Calendar.SECOND, 0);      // Detik 0

                        Calendar breakEnd = Calendar.getInstance(TimeZone.getTimeZone("Asia/Jakarta"));
                        breakEnd.setTime(workHoursSpecific.getDATE_WD());
                        breakEnd.set(Calendar.HOUR_OF_DAY, 12);   // Jam 12
                        breakEnd.set(Calendar.MINUTE, 40);        // Menit 40
                        breakEnd.set(Calendar.SECOND, 0);         // Detik 0

                        // Konversi ke zona waktu WIB jika diperlukan
                        TimeZone wibTimeZone = TimeZone.getTimeZone("Asia/Jakarta");
                        breakStart.setTimeZone(wibTimeZone);
                        breakEnd.setTimeZone(wibTimeZone);

                        // Debug log untuk memastikan jam sudah diatur dengan benar
                        System.out.println("Break Start: " + breakStart.getTime()); 
                        System.out.println("Break End: " + breakEnd.getTime());   

                        System.out.println("start time: " + startTimee.getTime()); 
                        System.out.println("end time: " + endTimee.getTime()); 

                        // Jika waktu mulai dan selesai melewati jam istirahat
                        if (startTimee.before(breakEnd) && endTimee.after(breakStart)) {
                            long overlapStart = Math.max(startTimee.getTimeInMillis(), breakStart.getTimeInMillis());
                            long overlapEnd = Math.min(endTimee.getTimeInMillis(), breakEnd.getTimeInMillis());
                            if (overlapStart < overlapEnd) {
                                long overlapDurationMillis = overlapEnd - overlapStart;
                                long overlapDurationMinutes = overlapDurationMillis / 60000;
                                shift1TotalTime = shift1TotalTime.subtract(new BigDecimal(overlapDurationMinutes));
                                System.out.println("Updated shift1TotalTime: " + shift1TotalTime); 
                            }
                        }
                    }
		            
                    workHoursSpecific.setSHIFT1_START_TIME(startTime);
                    workHoursSpecific.setSHIFT1_END_TIME(endTime);
                    workHoursSpecific.setSHIFT1_TOTAL_TIME(shift1TotalTime);
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
            
            System.out.println("Shift " + shift + " updated successfully.");
            return Optional.of(workHoursSpecific);
        } else {
            System.out.println("No WorkDayHoursSpecific record found for date: " + parsedDate);
            return Optional.empty();
        }
    }
    
//    private void updateOffAndSemiOff(DWorkDayHoursSpesific dWorkDayHoursSpecific, String dateWd, String description) {
//    	DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//        LocalDate localDate = LocalDate.parse(dateWd, formatter);
//        String nextDateStr = localDate.plusDays(1).format(formatter);
//    	
//    	Optional<WorkDay> workDayOpt = workDayRepo.findByDDateWd(dateWd);
//        List<DWorkDayHoursSpesific> workHoursSpecificOpt = dWorkDayHoursSpecificRepo.findDWdHoursListByDate(dateWd);
//        List<DWorkDayHoursSpesific> workHoursSpecificOpttomorrow = 
//        	    dWorkDayHoursSpecificRepo.findDWdHoursListByDate(nextDateStr);
//
//        workDayOpt.ifPresentOrElse(workDay -> {
//            // Ambil total waktu dari setiap shift
//            BigDecimal shift1 = dWorkDayHoursSpecific.getSHIFT1_TOTAL_TIME() != null ? dWorkDayHoursSpecific.getSHIFT1_TOTAL_TIME() : BigDecimal.ZERO;
//            BigDecimal shift2 = dWorkDayHoursSpecific.getSHIFT2_TOTAL_TIME() != null ? dWorkDayHoursSpecific.getSHIFT2_TOTAL_TIME() : BigDecimal.ZERO;
//            BigDecimal shift3 = dWorkDayHoursSpecific.getSHIFT3_TOTAL_TIME() != null ? dWorkDayHoursSpecific.getSHIFT3_TOTAL_TIME() : BigDecimal.ZERO;
//
//            // Log nilai shift
//            System.out.println("Shift1: " + shift1);
//            System.out.println("Shift2: " + shift2);
//            System.out.println("Shift3: " + shift3);
//
//            DayOfWeek dayOfWeek = getDayOfWeekFromDate(dateWd);
//
//            // Apply logic based on the description parameter
//            if ("WD_NORMAL".equals(description)) {
//                // Set IWD_SHIFT_1, IWD_SHIFT_2, IWD_SHIFT_3 based on shift times
//                workDay.setIWD_SHIFT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                workDay.setIWD_SHIFT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                workDay.setIWD_SHIFT_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//
//                System.out.println("IWD_SHIFT_1: " + workDay.getIWD_SHIFT_1());
//                System.out.println("IWD_SHIFT_2: " + workDay.getIWD_SHIFT_2());
//                System.out.println("IWD_SHIFT_3: " + workDay.getIWD_SHIFT_3());
//            } else if ("OT_TT".equals(description)) {
//                // Apply logic for OT_TT (Overtime Total Time)
//                workDay.setIOT_TT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                workDay.setIOT_TT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                workDay.setIOT_TT_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//
//                System.out.println("IOT_TT_1: " + workDay.getIOT_TT_1());
//                System.out.println("IOT_TT_2: " + workDay.getIOT_TT_2());
//                System.out.println("IOT_TT_3: " + workDay.getIOT_TT_3());
//            } else if ("OT_TL".equals(description)) {
//                // Apply logic for OT_TL (Overtime Total Leave)
//                workDay.setIOT_TL_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                workDay.setIOT_TL_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                workDay.setIOT_TL_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//
//                System.out.println("IOT_TL_1: " + workDay.getIOT_TL_1());
//                System.out.println("IOT_TL_2: " + workDay.getIOT_TL_2());
//                System.out.println("IOT_TL_3: " + workDay.getIOT_TL_3());
//            }
//
//            // Log and update OFF and SEMI_OFF based on the day of the week and shift times
//            if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
//                workDay.setOFF(BigDecimal.ONE);
//                workDay.setSEMI_OFF(BigDecimal.ZERO);
//            } else if (dayOfWeek == DayOfWeek.MONDAY) {
//                if (workDay.getIWD_SHIFT_1().compareTo(BigDecimal.ZERO) > 0 && workDay.getIWD_SHIFT_2().compareTo(BigDecimal.ZERO) > 0) {
//                    workDay.setOFF(BigDecimal.ZERO);
//                    workDay.setSEMI_OFF(BigDecimal.ZERO);
//                } else if (workDay.getIWD_SHIFT_1().compareTo(BigDecimal.ZERO) == 0 && workDay.getIWD_SHIFT_2().compareTo(BigDecimal.ZERO) == 0) {
//                    workDay.setOFF(BigDecimal.ONE);
//                    workDay.setSEMI_OFF(BigDecimal.ZERO);
//                } else {
//                    workDay.setOFF(BigDecimal.ZERO);
//                    workDay.setSEMI_OFF(BigDecimal.ONE);
//                }
//            } else if (dayOfWeek.getValue() >= DayOfWeek.TUESDAY.getValue() && dayOfWeek.getValue() <= DayOfWeek.FRIDAY.getValue()) {
//            	System.out.println("masuk sini 1");
//                if (workDay.getIWD_SHIFT_1().compareTo(BigDecimal.ZERO) == 0 && workDay.getIWD_SHIFT_2().compareTo(BigDecimal.ZERO) == 0 && workDay.getIWD_SHIFT_3().compareTo(BigDecimal.ZERO) == 0) {
//                    workDay.setOFF(BigDecimal.ONE);
//                    workDay.setSEMI_OFF(BigDecimal.ZERO);
//                	System.out.println("masuk sini 2");
//                } else if (workDay.getIWD_SHIFT_1().compareTo(BigDecimal.ZERO) > 0 && workDay.getIWD_SHIFT_2().compareTo(BigDecimal.ZERO) > 0 && workDay.getIWD_SHIFT_3().compareTo(BigDecimal.ZERO) > 0) {
//                    workDay.setOFF(BigDecimal.ZERO);
//                    workDay.setSEMI_OFF(BigDecimal.ZERO);
//                	System.out.println("masuk sini 3");
//                } else {
//                    workDay.setOFF(BigDecimal.ZERO);
//                    workDay.setSEMI_OFF(BigDecimal.ONE);
//                	System.out.println("masuk sini 4");
//
//                }
//            }
//
//            System.out.println("SEMI_OFF value: " + workDay.getSEMI_OFF());
//            System.out.println("OFF value: " + workDay.getOFF());
//
//            workDayRepo.save(workDay);
//        }, () -> {
//            throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
//        });
//    }

//    private void updateOffAndSemiOff(DWorkDayHoursSpesific dWorkDayHoursSpecific, String dateWd, String description) {
//        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//        LocalDate localDate = LocalDate.parse(dateWd, formatter);
//        String nextDateStr = localDate.plusDays(1).format(formatter);
//
//        Optional<WorkDay> workDayOpt = workDayRepo.findByDDateWd(dateWd);
//        List<DWorkDayHoursSpesific> workHoursSpecificOpt = dWorkDayHoursSpecificRepo.findDWdHoursListByDate(dateWd);
//        List<DWorkDayHoursSpesific> workHoursSpecificOpttomorrow = 
//        		dWorkDayHoursSpecificRepo.findDWdHoursListByDate(nextDateStr);
//
//        workDayOpt.ifPresentOrElse(workDay -> {
//            // Looping melalui data dari workHoursSpecificOpt untuk memeriksa setiap shift
//            for (DWorkDayHoursSpesific workHours : workHoursSpecificOpt) {
//                // Ambil total waktu dari setiap shift pada hari ini
//                BigDecimal shift1 = workHours.getSHIFT1_TOTAL_TIME() != null ? workHours.getSHIFT1_TOTAL_TIME() : BigDecimal.ZERO;
//                BigDecimal shift2 = workHours.getSHIFT2_TOTAL_TIME() != null ? workHours.getSHIFT2_TOTAL_TIME() : BigDecimal.ZERO;
//                BigDecimal shift3 = BigDecimal.ZERO;
//
//                Optional<DWorkDayHoursSpesific> tomorrowShiftOpt = dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(dateWd, workHours.getDESCRIPTION());
//                if (tomorrowShiftOpt.isPresent()) {
//                    shift3 = tomorrowShiftOpt.get().getSHIFT3_TOTAL_TIME() != null ? tomorrowShiftOpt.get().getSHIFT3_TOTAL_TIME() : BigDecimal.ZERO;
//                }
//                
//                // Log nilai shift
//                System.out.println("Shift1: " + shift1);
//                System.out.println("Shift2: " + shift2);
//                System.out.println("Shift3 (from tomorrow): " + shift3);
//
//                DayOfWeek dayOfWeek = getDayOfWeekFromDate(dateWd);
//
//                // Apply logic based on the description parameter
//                if ("WD_NORMAL".equals(workHours.getDESCRIPTION())) {
//                    // Logika untuk shift1, shift2, shift3
//                    workDay.setIWD_SHIFT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                    workDay.setIWD_SHIFT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                    workDay.setIWD_SHIFT_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//
//                    System.out.println("IWD_SHIFT_1: " + workDay.getIWD_SHIFT_1());
//                    System.out.println("IWD_SHIFT_2: " + workDay.getIWD_SHIFT_2());
//                    System.out.println("IWD_SHIFT_3: " + workDay.getIWD_SHIFT_3());
//                    
//                 // Logika untuk menentukan OFF dan SEMI_OFF
//                    if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
//                        // Sabtu dan Minggu OFF = 1, SEMI_OFF = 0
//                        workDay.setOFF(BigDecimal.ONE);
//                        workDay.setSEMI_OFF(BigDecimal.ZERO);
//                    } else if (dayOfWeek == DayOfWeek.MONDAY) {
//                        // Hari Senin OFF dan SEMI_OFF berdasarkan shift1, shift2, shift3
//                        if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0 && shift3.compareTo(BigDecimal.ZERO) == 0) {
//                            workDay.setOFF(BigDecimal.ONE);
//                            workDay.setSEMI_OFF(BigDecimal.ZERO);
//                        } else if ((shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) > 0) || 
//                                   (shift1.compareTo(BigDecimal.ZERO) > 0 && shift2.compareTo(BigDecimal.ZERO) == 0) || 
//                                   (shift3.compareTo(BigDecimal.ZERO) > 0)) {
//                            workDay.setOFF(BigDecimal.ZERO);
//                            workDay.setSEMI_OFF(BigDecimal.ONE);
//                        } else {
//                            workDay.setOFF(BigDecimal.ZERO);
//                            workDay.setSEMI_OFF(BigDecimal.ZERO);
//                        }
//                    } else if (dayOfWeek.getValue() >= DayOfWeek.TUESDAY.getValue() && dayOfWeek.getValue() <= DayOfWeek.FRIDAY.getValue()) {
//                        // Hari Selasa - Jumat OFF dan SEMI_OFF berdasarkan shift1, shift2, shift3
//                        if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0 && shift3.compareTo(BigDecimal.ZERO) == 0) {
//                            workDay.setOFF(BigDecimal.ONE);
//                            workDay.setSEMI_OFF(BigDecimal.ZERO);
//                        } else if ((shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) > 0) || 
//                                   (shift1.compareTo(BigDecimal.ZERO) > 0 && shift2.compareTo(BigDecimal.ZERO) == 0) || 
//                                   (shift3.compareTo(BigDecimal.ZERO) > 0)) {
//                            workDay.setOFF(BigDecimal.ZERO);
//                            workDay.setSEMI_OFF(BigDecimal.ONE);
//                        } else {
//                            workDay.setOFF(BigDecimal.ZERO);
//                            workDay.setSEMI_OFF(BigDecimal.ZERO);
//                        }
//                    }
//                }else if ("OT_TT".equals(workHours.getDESCRIPTION())) {
//                    // Apply logic for OT_TT (Overtime Total Time)
//                    workDay.setIOT_TT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                    workDay.setIOT_TT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                    workDay.setIOT_TT_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//
//                    System.out.println("IOT_TT_1: " + workDay.getIOT_TT_1());
//                    System.out.println("IOT_TT_2: " + workDay.getIOT_TT_2());
//                    System.out.println("IOT_TT_3: " + workDay.getIOT_TT_3());
//                } else if ("OT_TL".equals(workHours.getDESCRIPTION())) {
//                    // Apply logic for OT_TL (Overtime Total Leave)
//                    workDay.setIOT_TL_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                    workDay.setIOT_TL_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//                    workDay.setIOT_TL_3(shift3.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
//
//                    System.out.println("IOT_TL_1: " + workDay.getIOT_TL_1());
//                    System.out.println("IOT_TL_2: " + workDay.getIOT_TL_2());
//                    System.out.println("IOT_TL_3: " + workDay.getIOT_TL_3());
//                }
//
//                System.out.println("SEMI_OFF value: " + workDay.getSEMI_OFF());
//                System.out.println("OFF value: " + workDay.getOFF());
//            }
//
//            workDayRepo.save(workDay);
//        }, () -> {
//            throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
//        });
//    }

    
    private DayOfWeek getDayOfWeekFromDate(String dateWd) {
        return LocalDate.parse(dateWd, DateTimeFormatter.ofPattern("dd-MM-yyyy")).getDayOfWeek();
    }

    private void updateOffAndSemiOff(DWorkDayHoursSpesific dWorkDayHoursSpecific, String dateWd, String description) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
        LocalDate localDate = LocalDate.parse(dateWd, formatter);
        String prevDateStr = localDate.minusDays(1).format(formatter);
        String nextDateStr = localDate.plusDays(1).format(formatter);

        // Data untuk hari sebelumnya, hari ini, dan besok
        List<DWorkDayHoursSpesific> workHoursSpecificYesterday = dWorkDayHoursSpecificRepo.findDWdHoursListByDate(prevDateStr);
        List<DWorkDayHoursSpesific> workHoursSpecificOptnow = dWorkDayHoursSpecificRepo.findDWdHoursListByDate(dateWd);
        List<DWorkDayHoursSpesific> workHoursSpecificOpttomorrow = dWorkDayHoursSpecificRepo.findDWdHoursListByDate(nextDateStr);

        // Mengambil workDay untuk hari sebelumnya
        Optional<WorkDay> prevWorkDayOpt = workDayRepo.findByDDateWd(prevDateStr);
        // Mengambil workDay untuk hari ini
        Optional<WorkDay> workDayOpt = workDayRepo.findByDDateWd(dateWd);

    	System.out.println(prevWorkDayOpt);

        // Proses logika untuk hari sebelumnya
        prevWorkDayOpt.ifPresent(prevWorkDay -> {
        	System.out.println("Masuk sini a");
        	WorkDay currentPrevWorkDay = prevWorkDayOpt.get();
        	System.out.println("Masuk sini b");

            for (DWorkDayHoursSpesific workHours : workHoursSpecificYesterday) {
                BigDecimal shift1 = workHours.getSHIFT1_TOTAL_TIME() != null ? workHours.getSHIFT1_TOTAL_TIME() : BigDecimal.ZERO;
                BigDecimal shift2 = workHours.getSHIFT2_TOTAL_TIME() != null ? workHours.getSHIFT2_TOTAL_TIME() : BigDecimal.ZERO;
                BigDecimal shift3 = BigDecimal.ZERO;
            	System.out.println("Masuk sini c");
                // Ambil shift3 dari hari ini
                Optional<DWorkDayHoursSpesific> todayShiftOpt = dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(dateWd, workHours.getDESCRIPTION());
                if (todayShiftOpt.isPresent()) {
                    shift3 = todayShiftOpt.get().getSHIFT3_TOTAL_TIME() != null ? todayShiftOpt.get().getSHIFT3_TOTAL_TIME() : BigDecimal.ZERO;
                }

                DayOfWeek prevDayOfWeek = getDayOfWeekFromDate(prevDateStr);

                // Logika untuk hari sebelumnya
                if ("WD_NORMAL".equals(workHours.getDESCRIPTION())) {
                	currentPrevWorkDay.setIWD_SHIFT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                	currentPrevWorkDay.setIWD_SHIFT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                	currentPrevWorkDay.setIWD_SHIFT_3(workHours.getSHIFT3_TOTAL_TIME().compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                    // Tentukan OFF dan SEMI_OFF
                    if (prevDayOfWeek == DayOfWeek.SATURDAY || prevDayOfWeek == DayOfWeek.SUNDAY) {
                    	currentPrevWorkDay.setOFF(BigDecimal.ONE);
                    	currentPrevWorkDay.setSEMI_OFF(BigDecimal.ZERO);
                    } else {
                        // Logika OFF dan SEMI_OFF lainnya
                    	System.out.println("Shift 3 prev :" + shift3);
                        determineOffAndSemiOff(currentPrevWorkDay, prevDayOfWeek, shift1, shift2, shift3);
                    }
                }else if ("OT_TT".equals(workHours.getDESCRIPTION())) {
                    // Apply logic for OT_TT (Overtime Total Time)
                	currentPrevWorkDay.setIOT_TT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                	currentPrevWorkDay.setIOT_TT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                	currentPrevWorkDay.setIOT_TT_3(workHours.getSHIFT3_TOTAL_TIME().compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                    System.out.println("IOT_TT_1: " + currentPrevWorkDay.getIOT_TT_1());
                    System.out.println("IOT_TT_2: " + currentPrevWorkDay.getIOT_TT_2());
                    System.out.println("IOT_TT_3: " + currentPrevWorkDay.getIOT_TT_3());
                } else if ("OT_TL".equals(workHours.getDESCRIPTION())) {
                    // Apply logic for OT_TL (Overtime Total Leave)
                	currentPrevWorkDay.setIOT_TL_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                	currentPrevWorkDay.setIOT_TL_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                	currentPrevWorkDay.setIOT_TL_3(workHours.getSHIFT3_TOTAL_TIME().compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                    System.out.println("IOT_TL_1: " + currentPrevWorkDay.getIOT_TL_1());
                    System.out.println("IOT_TL_2: " + currentPrevWorkDay.getIOT_TL_2());
                    System.out.println("IOT_TL_3: " + currentPrevWorkDay.getIOT_TL_3());
                }
            }
        	System.out.println("Masuk sini d");
            workDayRepo.save(currentPrevWorkDay);
        });

        // Proses logika untuk hari ini
        workDayOpt.ifPresentOrElse(workDay -> {
            for (DWorkDayHoursSpesific workHours : workHoursSpecificOptnow) {
                BigDecimal shift1 = workHours.getSHIFT1_TOTAL_TIME() != null ? workHours.getSHIFT1_TOTAL_TIME() : BigDecimal.ZERO;
                BigDecimal shift2 = workHours.getSHIFT2_TOTAL_TIME() != null ? workHours.getSHIFT2_TOTAL_TIME() : BigDecimal.ZERO;
                BigDecimal shift3 = BigDecimal.ZERO;

                Optional<DWorkDayHoursSpesific> tomorrowShiftOpt = dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(nextDateStr, workHours.getDESCRIPTION());

                if (tomorrowShiftOpt.isPresent()) {
                    shift3 = tomorrowShiftOpt.get().getSHIFT3_TOTAL_TIME() != null ? tomorrowShiftOpt.get().getSHIFT3_TOTAL_TIME() : BigDecimal.ZERO;
                }

                DayOfWeek todayDayOfWeek = getDayOfWeekFromDate(dateWd);

                if ("WD_NORMAL".equals(workHours.getDESCRIPTION())) {
                    workDay.setIWD_SHIFT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                    workDay.setIWD_SHIFT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                    workDay.setIWD_SHIFT_3(workHours.getSHIFT3_TOTAL_TIME().compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                    // Tentukan OFF dan SEMI_OFF
                    if (todayDayOfWeek == DayOfWeek.SATURDAY || todayDayOfWeek == DayOfWeek.SUNDAY) {
                        workDay.setOFF(BigDecimal.ONE);
                        workDay.setSEMI_OFF(BigDecimal.ZERO);
                    } else {
                        // Logika OFF dan SEMI_OFF lainnya
                    	System.out.println("Shift 3 now :" + shift3);
                        determineOffAndSemiOff(workDay, todayDayOfWeek, shift1, shift2, shift3);
                    }
                }else if ("OT_TT".equals(workHours.getDESCRIPTION())) {
                    // Apply logic for OT_TT (Overtime Total Time)
                    workDay.setIOT_TT_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                    workDay.setIOT_TT_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                    workDay.setIOT_TT_3(workHours.getSHIFT3_TOTAL_TIME().compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                    System.out.println("IOT_TT_1: " + workDay.getIOT_TT_1());
                    System.out.println("IOT_TT_2: " + workDay.getIOT_TT_2());
                    System.out.println("IOT_TT_3: " + workDay.getIOT_TT_3());
                } else if ("OT_TL".equals(workHours.getDESCRIPTION())) {
                    // Apply logic for OT_TL (Overtime Total Leave)
                    workDay.setIOT_TL_1(shift1.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                    workDay.setIOT_TL_2(shift2.compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);
                    workDay.setIOT_TL_3(workHours.getSHIFT3_TOTAL_TIME().compareTo(BigDecimal.ZERO) > 0 ? BigDecimal.ONE : BigDecimal.ZERO);

                    System.out.println("IOT_TL_1: " + workDay.getIOT_TL_1());
                    System.out.println("IOT_TL_2: " + workDay.getIOT_TL_2());
                    System.out.println("IOT_TL_3: " + workDay.getIOT_TL_3());
                }
            }
            workDayRepo.save(workDay);
        }, () -> {
            throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
        });
    }

    private void determineOffAndSemiOff(WorkDay workDay, DayOfWeek dayOfWeek, BigDecimal shift1, BigDecimal shift2, BigDecimal shift3) {
            if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0 && shift3.compareTo(BigDecimal.ZERO) == 0) {
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else if ((shift1.compareTo(BigDecimal.ZERO) > 0 || 
                    shift2.compareTo(BigDecimal.ZERO) > 0 || 
                    shift3.compareTo(BigDecimal.ZERO) > 0) && 
                   !(shift1.compareTo(BigDecimal.ZERO) > 0 && 
                     shift2.compareTo(BigDecimal.ZERO) > 0 && 
                     shift3.compareTo(BigDecimal.ZERO) > 0)) {
              workDay.setOFF(BigDecimal.ZERO);
              workDay.setSEMI_OFF(BigDecimal.ONE);
          } else {
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            }
    }

}