package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.lang.reflect.Field;

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
import sri.sysint.sri_starter_back.repository.DWorkDayHoursSpecificRepo;
import sri.sysint.sri_starter_back.model.DWorkDayHoursSpesific;
import sri.sysint.sri_starter_back.service.DWorkDayHoursSpecificServiceImpl; 



import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.repository.WorkDayRepo;

@Service
@Transactional
public class WorkDayServiceImpl {

    @Autowired
    private WorkDayRepo workDayRepo;
	@Autowired
	private DWorkDayHoursSpecificRepo dWorkDayHoursSpecificRepo;
    @Autowired
    private DWorkDayHoursServiceImpl dWorkDayHoursServiceImpl; 
    @Autowired
    private DWorkDayHoursSpecificServiceImpl dWorkDayHoursSpecificServiceImpl; 

    public WorkDayServiceImpl(WorkDayRepo workDayRepo) {
        this.workDayRepo = workDayRepo;
    }
    
    public List<WorkDay> getAllWorkDays() {
        Iterable<WorkDay> workDays = workDayRepo.getDataOrderByDateWd();
        List<WorkDay> workDayList = new ArrayList<>();
        for (WorkDay item : workDays) {
            WorkDay workDayTemp = new WorkDay(item);
            workDayList.add(workDayTemp);
        }

        return workDayList;
    }
    
    public List<WorkDay> getAllWorkDaysByDateRange(Date startDate, Date endDate) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String formattedStartDate = dateFormat.format(startDate);
        String formattedEndDate = dateFormat.format(endDate);
        
        System.out.println("Fetching WorkDays between " + formattedStartDate + " and " + formattedEndDate);

        try {
            List<WorkDay> workDays = workDayRepo.findAllByDateRange(formattedStartDate, formattedEndDate);
            return workDays;
        } catch (Exception e) {
            System.err.println("Error fetching work days by date range: " + e.getMessage());
            throw e;
        }
    }

    public Optional<WorkDay> getWorkDayByDate(Date date) {
        
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String formattedDate = dateFormat.format(date);
        System.out.println("Formatted date being sent to repository: " + formattedDate);

        Optional<WorkDay> result = workDayRepo.findByDDateWd(formattedDate);
        if (result.isPresent()) {
            System.out.println("Data found: " + result.get());
        } else {
            System.out.println("No data found for date: " + formattedDate);
        }

        return result;
    }

    public WorkDay saveWorkDay(WorkDay workDay) {
        try {
            workDay.setCREATION_DATE(new Date());
            workDay.setLAST_UPDATE_DATE(new Date());
            workDay.setSTATUS(BigDecimal.valueOf(1));

            LocalDate dateWD;

            try {
                dateWD = LocalDate.parse(workDay.getDATE_WD().toString());
            } catch (Exception isoParseException) {
                try {
                    SimpleDateFormat fullDateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                    Date parsedDate = fullDateFormat.parse(workDay.getDATE_WD().toString());
                    dateWD = parsedDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                } catch (Exception parseException) {
                    throw new RuntimeException("Unable to parse DATE_WD: " + workDay.getDATE_WD(), parseException);
                }
            }

            updateOffAndSemiOffSU(workDay, dateWD);

            return workDayRepo.save(workDay);
        } catch (Exception e) {
            System.err.println("Error saving work day: " + e.getMessage());
            throw e;
        }
    }



    public WorkDay updateWorkDay(WorkDay workDay) throws Exception {
        try {
            Optional<WorkDay> currentWorkDayOpt = workDayRepo.findById(workDay.getDATE_WD());

            if (currentWorkDayOpt.isPresent()) {
                WorkDay currentWorkDay = currentWorkDayOpt.get();

                currentWorkDay.setIWD_SHIFT_1(workDay.getIWD_SHIFT_1());
                currentWorkDay.setIWD_SHIFT_2(workDay.getIWD_SHIFT_2());
                currentWorkDay.setIWD_SHIFT_3(workDay.getIWD_SHIFT_3());
                currentWorkDay.setIOT_TL_1(workDay.getIOT_TL_1());
                currentWorkDay.setIOT_TL_2(workDay.getIOT_TL_2());
                currentWorkDay.setIOT_TL_3(workDay.getIOT_TL_3());
                currentWorkDay.setIOT_TT_1(workDay.getIOT_TT_1());
                currentWorkDay.setIOT_TT_2(workDay.getIOT_TT_2());
                currentWorkDay.setIOT_TT_3(workDay.getIOT_TT_3());
                currentWorkDay.setOFF(workDay.getOFF());
                currentWorkDay.setSEMI_OFF(workDay.getSEMI_OFF());

                LocalDate dateWD;

                try {
                    dateWD = LocalDate.parse(workDay.getDATE_WD().toString());
                } catch (Exception isoParseException) {
                    try {
                        SimpleDateFormat fullDateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
                        Date parsedDate = fullDateFormat.parse(workDay.getDATE_WD().toString());
                        dateWD = parsedDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                    } catch (Exception e) {
                        throw new RuntimeException("Unable to parse DATE_WD: " + workDay.getDATE_WD(), e);
                    }
                }

                updateOffAndSemiOffSU(currentWorkDay, dateWD);                
                currentWorkDay.setLAST_UPDATE_DATE(new Date());
                currentWorkDay.setLAST_UPDATED_BY(workDay.getLAST_UPDATED_BY());

                workDayRepo.save(currentWorkDay);
                
                updateSpesificHour(currentWorkDay, workDay.getDATE_WD());
                
                return currentWorkDay;
            } else {
                throw new RuntimeException("WorkDay with date " + workDay.getDATE_WD() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating work day: " + e.getMessage());
            throw e;
        }
    }

    public WorkDay deleteWorkDay(WorkDay workDay) {
        try {
            Optional<WorkDay> currentWorkDayOpt = workDayRepo.findById(workDay.getDATE_WD());

            if (currentWorkDayOpt.isPresent()) {
                WorkDay currentWorkDay = currentWorkDayOpt.get();

                currentWorkDay.setSTATUS(BigDecimal.valueOf(0));
                currentWorkDay.setLAST_UPDATE_DATE(new Date());
                currentWorkDay.setLAST_UPDATED_BY(workDay.getLAST_UPDATED_BY());

                return workDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("WorkDay with date " + workDay.getDATE_WD() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error deleting work day: " + e.getMessage());
            throw e;
        }
    }

    public WorkDay restoreWorkDay(WorkDay workDay) {
        try {
            Optional<WorkDay> currentWorkDayOpt = workDayRepo.findById(workDay.getDATE_WD());

            if (currentWorkDayOpt.isPresent()) {
                WorkDay currentWorkDay = currentWorkDayOpt.get();

                currentWorkDay.setSTATUS(BigDecimal.valueOf(1)); 
                currentWorkDay.setLAST_UPDATE_DATE(new Date());
                currentWorkDay.setLAST_UPDATED_BY(workDay.getLAST_UPDATED_BY());

                return workDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("WorkDay with date " + workDay.getDATE_WD() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring work day: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllWorkDays() {
        workDayRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportWDsExcel() throws IOException {
        List<WorkDay> workDays = workDayRepo.getDataOrderByDateWd();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(workDays);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<WorkDay> workDays) throws IOException {
        List<String> headers = List.of(
            "DATE_WD",
            "WD_SHIFT_3",
            "WD_SHIFT_2",
            "WD_SHIFT_1",
            "OT_TL_3",
            "OT_TL_2",
            "OT_TL_1",
            "OT_TT_3",
            "OT_TT_2",
            "OT_TT_1",
            "OFF",
            "SEMI_OFF"
        );

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("WORKDAYS DATA");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Row headerRow = sheet.createRow(0);
            SimpleDateFormat dateFormat = new SimpleDateFormat("E MMMM dd, yyyy");

            for (int j = 0; j < headers.size(); j++) {
                Row dataRow = sheet.createRow(j); 
                
                Cell headerCell = dataRow.createCell(0);
                headerCell.setCellValue(headers.get(j));
                headerCell.setCellStyle(borderStyle);

                for (int i = 0; i < workDays.size(); i++) {
                    WorkDay workDay = workDays.get(i);
                    Cell dataCell = dataRow.createCell(i+1);
                    
                    switch (j) {
                        case 0: // DATE_WD
                            String formattedDate = dateFormat.format(workDay.getDATE_WD());
                            dataCell.setCellValue(formattedDate);
                            break;
                        case 1: // WD_SHIFT_3
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIWD_SHIFT_3()));
                            break;
                        case 2: // WD_SHIFT_2
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIWD_SHIFT_2()));
                            break;
                        case 3: // WD_SHIFT_1
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIWD_SHIFT_1()));
                            break;
                        case 4: // OT_TL_3
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIOT_TL_3()));
                            break;
                        case 5: // OT_TL_2
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIOT_TL_2()));
                            break;
                        case 6: // OT_TL_1
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIOT_TL_1()));
                            break;
                        case 7: // OT_TT_3
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIOT_TT_3()));
                            break;
                        case 8: // OT_TT_2
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIOT_TT_2()));
                            break;
                        case 9: // OT_TT_1
                            dataCell.setCellValue(convertToDisplayValue(workDay.getIOT_TT_1()));
                            break;
                        case 10: // OFF
                            dataCell.setCellValue(convertToDisplayValue(workDay.getOFF()));
                            break;
                        case 11: // SEMI_OFF
                            dataCell.setCellValue(convertToDisplayValue(workDay.getSEMI_OFF()));
                            break;
                    }
                    
                    dataCell.setCellStyle(borderStyle);
                }
            }

            for (int i = 0; i < headers.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export WorkDay data");
            throw e;
        } finally {
            out.close();
            workbook.close(); 
        }
    }


    private String convertToDisplayValue(BigDecimal value) {
        return (value != null && value.compareTo(BigDecimal.ONE) == 0) ? "v" : "";
    }
    
    public WorkDay turnOnOvertime(String dateWd) throws Exception {
        try {
            Optional<WorkDay> currentWorkDayOpt = workDayRepo.findByDDateWd(dateWd);
            String OT_TT_1 = ("OT_TT_1");
            String OT_TT_2 = ("OT_TT_2");
            String OT_TT_3 = ("OT_TT_3");
            
            String OT_TL_1 = ("OT_TL_1");
            String OT_TL_2 = ("OT_TL_2");
            String OT_TL_3 = ("OT_TL_3");
            
            if (currentWorkDayOpt.isPresent()) {
                WorkDay currentWorkDay = currentWorkDayOpt.get();

                currentWorkDay.setIWD_SHIFT_1(BigDecimal.ZERO);
                currentWorkDay.setIWD_SHIFT_2(BigDecimal.ZERO);
                currentWorkDay.setIWD_SHIFT_3(BigDecimal.ZERO);

                turnOnShift(dateWd, OT_TT_1);
                turnOnShift(dateWd, OT_TT_2);
                turnOnShift(dateWd, OT_TT_3);
                turnOnShift(dateWd, OT_TL_1);
                turnOnShift(dateWd, OT_TL_2);
                turnOnShift(dateWd, OT_TL_3);

                LocalDate parsedDateWD;

                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
                try {
                    parsedDateWD = LocalDate.parse(dateWd, formatter);
                } catch (DateTimeParseException e) {
                    throw new RuntimeException("Unable to parse dateWd: " + dateWd, e);
                }

                updateOffAndSemiOffSU(currentWorkDay, parsedDateWD);
                currentWorkDay.setLAST_UPDATE_DATE(new Date());

                return workDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error turning on overtime: " + e.getMessage());
            throw e;
        }
    }
   
    public WorkDay turnOffShift(String dateWd, String shift) throws Exception {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        Date parsedDate = dateFormat.parse(dateWd); 
        
        WorkDay updatedWorkDay = updateShiftValue(dateWd, shift, BigDecimal.ZERO);
        return updatedWorkDay;
    }

    public WorkDay turnOnShift(String dateWd, String shift) throws Exception {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        Date parsedDate = dateFormat.parse(dateWd); 
        
        WorkDay updatedWorkDay = updateShiftValue(dateWd, shift, BigDecimal.ONE);
        return updatedWorkDay;
    }


    public WorkDay updateShiftValue(String dateWd, String shift, BigDecimal newValue) throws Exception {
        try {
            Optional<WorkDay> currentWorkDayOpt = workDayRepo.findByDDateWd(dateWd);

            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
            Date parsedDate = dateFormat.parse(dateWd); 
            String description;

            if (currentWorkDayOpt.isPresent()) {
                WorkDay currentWorkDay = currentWorkDayOpt.get();

                switch (shift.toUpperCase()) {
                    case "SHIFT 1":
                        currentWorkDay.setIWD_SHIFT_1(newValue);
                        description = "WD_NORMAL";
                        break;
                    case "SHIFT 2":
                        currentWorkDay.setIWD_SHIFT_2(newValue);
                        description = "WD_NORMAL";
                        break;
                    case "SHIFT 3":
                        currentWorkDay.setIWD_SHIFT_3(newValue);
                        description = "WD_NORMAL";
                        break;
                    case "OT_TL_1":
                        currentWorkDay.setIOT_TL_1(newValue);
                        description = "OT_TL";
                        break;
                    case "OT_TL_2":
                        currentWorkDay.setIOT_TL_2(newValue);
                        description = "OT_TL";
                        break;
                    case "OT_TL_3":
                        currentWorkDay.setIOT_TL_3(newValue);
                        description = "OT_TL";
                        break;
                    case "OT_TT_1":
                        currentWorkDay.setIOT_TT_1(newValue);
                        description = "OT_TT";
                        break;
                    case "OT_TT_2":
                        currentWorkDay.setIOT_TT_2(newValue);
                        description = "OT_TT";
                        break;
                    case "OT_TT_3":
                        currentWorkDay.setIOT_TT_3(newValue);
                        description = "OT_TT";
                        break;
                    default:
                        throw new IllegalArgumentException("Invalid shift specified: " + shift);
                }
                updateOffAndSemiOff(currentWorkDay, dateWd);

                return workDayRepo.save(currentWorkDay);
            } else {
                throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
            }
        } catch (ParseException e) {
            System.err.println("Error parsing date: " + e.getMessage());
            throw new RuntimeException("Invalid date format provided: " + dateWd, e);  
        } catch (Exception e) {
            System.err.println("Error updating shift: " + e.getMessage());
            throw e;  
        }
    }
    
    private void updateOffAndSemiOffSU(WorkDay workDay, LocalDate dateWd) {
        BigDecimal shift1 = workDay.getIWD_SHIFT_1();
        BigDecimal shift2 = workDay.getIWD_SHIFT_2();
        BigDecimal shift3 = workDay.getIWD_SHIFT_3();

        DayOfWeek dayOfWeek = dateWd.getDayOfWeek();

        // Logic for Saturday and Sunday
        if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
            workDay.setOFF(BigDecimal.ONE);
            workDay.setSEMI_OFF(BigDecimal.ZERO);
            return;
        }

        // Logic for Monday
        if (dayOfWeek == DayOfWeek.MONDAY) {
            if (shift1.compareTo(BigDecimal.ONE) == 0 && shift2.compareTo(BigDecimal.ONE) == 0) {
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0) {
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else {
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ONE);
            }
            return;
        }

        // Logic for Tuesday - Friday
        if (dayOfWeek.getValue() >= DayOfWeek.TUESDAY.getValue() && dayOfWeek.getValue() <= DayOfWeek.FRIDAY.getValue()) {
            if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0 && shift3.compareTo(BigDecimal.ZERO) == 0) {
                // All shifts inactive
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else if (shift1.compareTo(BigDecimal.ONE) == 0 && shift2.compareTo(BigDecimal.ONE) == 0 && shift3.compareTo(BigDecimal.ONE) == 0) {
                // All shifts active
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else {
                // Mixed active/inactive shifts
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ONE);
            }
        }
    }

    private void updateOffAndSemiOff(WorkDay workDay, String dateWd) throws Exception {
        BigDecimal shift1 = workDay.getIWD_SHIFT_1();
        BigDecimal shift2 = workDay.getIWD_SHIFT_2();
        BigDecimal shift3 = workDay.getIWD_SHIFT_3();

        DayOfWeek dayOfWeek = getDayOfWeekFromDate(dateWd);

        // Sabtu - Minggu
        if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
            workDay.setOFF(BigDecimal.ONE);
            workDay.setSEMI_OFF(BigDecimal.ZERO);
            return;
        }

        // Logika untuk Senin
        if (dayOfWeek == DayOfWeek.MONDAY) {
            if (shift1.compareTo(BigDecimal.ONE) == 0 && shift2.compareTo(BigDecimal.ONE) == 0) {
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0) {
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else {
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ONE);
            }
            return;
        }

        //  logika selasa - jumat 
        if (dayOfWeek.getValue() >= DayOfWeek.TUESDAY.getValue() && dayOfWeek.getValue() <= DayOfWeek.FRIDAY.getValue()) {
            if (shift1.compareTo(BigDecimal.ZERO) == 0 && shift2.compareTo(BigDecimal.ZERO) == 0 && shift3.compareTo(BigDecimal.ZERO) == 0) {
                // Semua shift non-aktif
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else if (shift1.compareTo(BigDecimal.ONE) == 0 && shift2.compareTo(BigDecimal.ONE) == 0 && shift3.compareTo(BigDecimal.ONE) == 0) {
                // Semua shift aktif
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
            } else {
                // Campuran aktif/non-aktif
                workDay.setOFF(BigDecimal.ZERO);
                workDay.setSEMI_OFF(BigDecimal.ONE);
            }
        }
        
      updateSpesificHour(workDay, workDay.getDATE_WD());

    }


    private DayOfWeek getDayOfWeekFromDate(String dateWd) {
        // Mengubah string dateWd ke LocalDate dan mendapatkan hari dalam minggu
        LocalDate date = LocalDate.parse(dateWd, DateTimeFormatter.ofPattern("dd-MM-yyyy"));
        return date.getDayOfWeek();
    }
    
    private void updateSpesificHour(WorkDay workDay, Date dateWd) throws Exception {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String formattedDate = dateFormat.format(dateWd);

        Map<String, String> shiftDescriptionMap = Map.of(
            "IWD_SHIFT_1", "WD_NORMAL",
            "IWD_SHIFT_2", "WD_NORMAL",
            "IWD_SHIFT_3", "WD_NORMAL",
            "IOT_TL_1", "OT_TL",
            "IOT_TL_2", "OT_TL",
            "IOT_TL_3", "OT_TL",
            "IOT_TT_1", "OT_TT",
            "IOT_TT_2", "OT_TT",
            "IOT_TT_3", "OT_TT"
        );

        for (Map.Entry<String, String> entry : shiftDescriptionMap.entrySet()) {
            String shiftFieldName = entry.getKey();
            String description = entry.getValue();

            try {
                Field shiftField = WorkDay.class.getDeclaredField(shiftFieldName);
                shiftField.setAccessible(true);
                BigDecimal shiftValue = (BigDecimal) shiftField.get(workDay);

                String startTime = "00:00";
                String endTime = "00:00";
                if (shiftValue != null && shiftValue.compareTo(BigDecimal.ONE) == 0) {
                    startTime = "07:00";
                    endTime = "15:00";
                }

                Optional<DWorkDayHoursSpesific> currentWorkHoursSpecificOpt =
                    dWorkDayHoursSpecificRepo.findDWdHoursByDateAndDescription(formattedDate, description);

                if (currentWorkHoursSpecificOpt.isPresent()) {
                	System.out.println("Masuk update hours spc");
                    DWorkDayHoursSpesific currentWorkHoursSpecific = currentWorkHoursSpecificOpt.get();

                    dWorkDayHoursSpecificServiceImpl.updateShiftTimesforWd(startTime, endTime, formattedDate, description, Integer.parseInt(shiftFieldName.split("_")[2]));
                }
            } catch (NoSuchFieldException | IllegalAccessException e) {
                throw new Exception("Error accessing field: " + shiftFieldName, e);
            }
        }
    }
}
