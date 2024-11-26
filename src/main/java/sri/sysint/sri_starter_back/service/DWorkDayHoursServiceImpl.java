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

import sri.sysint.sri_starter_back.model.DWorkDayHours;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.repository.DWorkDayHoursRepo;
import sri.sysint.sri_starter_back.repository.WorkDayRepo;

@Service
@Transactional
public class DWorkDayHoursServiceImpl {

    @Autowired
    private DWorkDayHoursRepo dWorkDayHoursRepo;
    
    
    @Autowired
    private WorkDayRepo workDayRepo;

    public DWorkDayHoursServiceImpl(DWorkDayHoursRepo dWorkDayHoursRepo) {
        this.dWorkDayHoursRepo = dWorkDayHoursRepo;
    }
    
    public BigDecimal getNewId() {
        return dWorkDayHoursRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<DWorkDayHours> getAllWorkDayHours() {
        Iterable<DWorkDayHours> workHours = dWorkDayHoursRepo.findAll();
        List<DWorkDayHours> workHourList = new ArrayList<>();
        for (DWorkDayHours item : workHours) {
            workHourList.add(new DWorkDayHours(item));
        }
        return workHourList;
    }

    public Optional<DWorkDayHours> getWorkDayHoursByDate(Date date) {
        
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String formattedDate = dateFormat.format(date);
        System.out.println("Formatted date being sent to repository: " + formattedDate);

        Optional<DWorkDayHours> result = dWorkDayHoursRepo.findDWdHoursByDate(formattedDate);
        if (result.isPresent()) {
            System.out.println("Data found: " + result.get());
        } else {
            System.out.println("No data found for date: " + formattedDate);
        }
        
        return result;
        
    }
    
    public Optional<DWorkDayHours> getWorkDayHoursByDateDesc(Date date, String description) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String formattedDate = dateFormat.format(date);
        System.out.println("Formatted date being sent to repository: " + formattedDate);
        
        Optional<DWorkDayHours> result = dWorkDayHoursRepo.findDWdHoursByDateAndDescription(formattedDate, description);
        
        if (result.isPresent()) {
            System.out.println("Data found: " + result.get());
        } else {
            System.out.println("No data found for date: " + formattedDate + " and description: " + description);
        }
        
        return result;
    }
    
    public DWorkDayHours saveWorkDayHours(DWorkDayHours workHours) {
        try {
            // Ensure we have the correct values set before saving
            workHours.setDETAIL_WD_HOURS_ID(getNewId());
            workHours.setCREATION_DATE(new Date());
            workHours.setLAST_UPDATE_DATE(new Date());
            workHours.setSTATUS(BigDecimal.valueOf(1)); 

            // Save without additional manipulation
            return dWorkDayHoursRepo.save(workHours);
        } catch (Exception e) {
            System.err.println("Error saving work hours: " + e.getMessage());
            throw e;
        }
    }

    public DWorkDayHours updateWorkDayHours(DWorkDayHours workHours) {
        try {
            Optional<DWorkDayHours> currentWorkHoursOpt = dWorkDayHoursRepo.findById(workHours.getDETAIL_WD_HOURS_ID());

            if (currentWorkHoursOpt.isPresent()) {
                DWorkDayHours currentWorkHours = currentWorkHoursOpt.get();

                // Update relevant fields (HOUR_1 to HOUR_24)
//                currentWorkHours.setDATE_WD(workHours.getDATE_WD());
                currentWorkHours.setHOUR_1(workHours.getHOUR_1());
                currentWorkHours.setHOUR_2(workHours.getHOUR_2());
                currentWorkHours.setHOUR_3(workHours.getHOUR_3());
                currentWorkHours.setHOUR_4(workHours.getHOUR_4());
                currentWorkHours.setHOUR_5(workHours.getHOUR_5());
                currentWorkHours.setHOUR_6(workHours.getHOUR_6());
                currentWorkHours.setHOUR_7(workHours.getHOUR_7());
                currentWorkHours.setHOUR_8(workHours.getHOUR_8());
                currentWorkHours.setHOUR_9(workHours.getHOUR_9());
                currentWorkHours.setHOUR_10(workHours.getHOUR_10());
                currentWorkHours.setHOUR_11(workHours.getHOUR_11());
                currentWorkHours.setHOUR_12(workHours.getHOUR_12());
                currentWorkHours.setHOUR_13(workHours.getHOUR_13());
                currentWorkHours.setHOUR_14(workHours.getHOUR_14());
                currentWorkHours.setHOUR_15(workHours.getHOUR_15());
                currentWorkHours.setHOUR_16(workHours.getHOUR_16());
                currentWorkHours.setHOUR_17(workHours.getHOUR_17());
                currentWorkHours.setHOUR_18(workHours.getHOUR_18());
                currentWorkHours.setHOUR_19(workHours.getHOUR_19());
                currentWorkHours.setHOUR_20(workHours.getHOUR_20());
                currentWorkHours.setHOUR_21(workHours.getHOUR_21());
                currentWorkHours.setHOUR_22(workHours.getHOUR_22());
                currentWorkHours.setHOUR_23(workHours.getHOUR_23());
                currentWorkHours.setHOUR_24(workHours.getHOUR_24());
                currentWorkHours.setDESCRIPTION(workHours.getDESCRIPTION());

                currentWorkHours.setLAST_UPDATE_DATE(new Date());
                currentWorkHours.setLAST_UPDATED_BY(workHours.getLAST_UPDATED_BY());

                return dWorkDayHoursRepo.save(currentWorkHours);
            } else {
                throw new RuntimeException("WorkHours with date " + workHours.getDATE_WD() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating work hours: " + e.getMessage());
            throw e;
        }
    }


    // Delete specific work hours (soft delete by changing status)
    public DWorkDayHours deleteWorkDayHours(DWorkDayHours workHours) {
        try {
            Optional<DWorkDayHours> currentWorkHoursOpt = dWorkDayHoursRepo.findById(workHours.getDETAIL_WD_HOURS_ID());

            if (currentWorkHoursOpt.isPresent()) {
                DWorkDayHours currentWorkHours = currentWorkHoursOpt.get();

                currentWorkHours.setSTATUS(BigDecimal.valueOf(0)); 
                currentWorkHours.setLAST_UPDATE_DATE(new Date());
                currentWorkHours.setLAST_UPDATED_BY(workHours.getLAST_UPDATED_BY());

                return dWorkDayHoursRepo.save(currentWorkHours);
            } else {
                throw new RuntimeException("WorkHours with date " + workHours.getDATE_WD() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error deleting work hours: " + e.getMessage());
            throw e;
        }
    }

    // Restore work hours (soft undelete by changing status)
    public DWorkDayHours restoreWorkDayHours(DWorkDayHours workHours) {
        try {
            Optional<DWorkDayHours> currentWorkHoursOpt = dWorkDayHoursRepo.findById(workHours.getDETAIL_WD_HOURS_ID());

            if (currentWorkHoursOpt.isPresent()) {
                DWorkDayHours currentWorkHours = currentWorkHoursOpt.get();

                currentWorkHours.setSTATUS(BigDecimal.valueOf(1)); 
                currentWorkHours.setLAST_UPDATE_DATE(new Date());
                currentWorkHours.setLAST_UPDATED_BY(workHours.getLAST_UPDATED_BY());

                return dWorkDayHoursRepo.save(currentWorkHours);
            } else {
                throw new RuntimeException("WorkHours with date " + workHours.getDATE_WD() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring work hours: " + e.getMessage());
            throw e;
        }
    }

    // Delete all work hours records
    public void deleteAllWorkDayHours() {
        dWorkDayHoursRepo.deleteAll();
    }

    // Export to Excel function
    public ByteArrayInputStream exportWorkDayHoursExcel() throws IOException {
        List<DWorkDayHours> workHoursList = dWorkDayHoursRepo.findAll();
        return dataToExcel(workHoursList);
    }

    private ByteArrayInputStream dataToExcel(List<DWorkDayHours> workHoursList) throws IOException {
        List<String> columns = List.of(
            "No.", "DETAIL_WD_HOURS_ID", "DATE_D_WH", "HOUR_1", "HOUR_2", "HOUR_3", "HOUR_4", "HOUR_5", 
            "HOUR_6", "HOUR_7", "HOUR_8", "HOUR_9", "HOUR_10", "HOUR_11", "HOUR_12", 
            "HOUR_13", "HOUR_14", "HOUR_15", "HOUR_16", "HOUR_17", "HOUR_18", 
            "HOUR_19", "HOUR_20", "HOUR_21", "HOUR_22", "HOUR_23", "HOUR_24", "DESCRIPTION"
        );

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("WorkDayHours");

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
            for (int i = 0; i < workHoursList.size(); i++) {
                DWorkDayHours workHours = workHoursList.get(i);
                Row dataRow = sheet.createRow(i + 1);

                // "No." column - Sequential number
                Cell noCell = dataRow.createCell(0);
                noCell.setCellValue(i + 1); 
                noCell.setCellStyle(borderStyle);

                // DETAIL_WD_HOURS_ID
                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(workHours.getDETAIL_WD_HOURS_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                // DATE_D_WH
                Cell dateCell = dataRow.createCell(2);
                String formattedDate = dateFormat.format(workHours.getDATE_WD());
                dateCell.setCellValue(formattedDate);
                dateCell.setCellStyle(borderStyle);

                // Populate other cells with hour data
                dataRow.createCell(3).setCellValue(convertToDisplayValue(workHours.getHOUR_1()));
                dataRow.createCell(4).setCellValue(convertToDisplayValue(workHours.getHOUR_2()));
                dataRow.createCell(5).setCellValue(convertToDisplayValue(workHours.getHOUR_3()));
                dataRow.createCell(6).setCellValue(convertToDisplayValue(workHours.getHOUR_4()));
                dataRow.createCell(7).setCellValue(convertToDisplayValue(workHours.getHOUR_5()));
                dataRow.createCell(8).setCellValue(convertToDisplayValue(workHours.getHOUR_6()));
                dataRow.createCell(9).setCellValue(convertToDisplayValue(workHours.getHOUR_7()));
                dataRow.createCell(10).setCellValue(convertToDisplayValue(workHours.getHOUR_8()));
                dataRow.createCell(11).setCellValue(convertToDisplayValue(workHours.getHOUR_9()));
                dataRow.createCell(12).setCellValue(convertToDisplayValue(workHours.getHOUR_10()));
                dataRow.createCell(13).setCellValue(convertToDisplayValue(workHours.getHOUR_11()));
                dataRow.createCell(14).setCellValue(convertToDisplayValue(workHours.getHOUR_12()));
                dataRow.createCell(15).setCellValue(convertToDisplayValue(workHours.getHOUR_13()));
                dataRow.createCell(16).setCellValue(convertToDisplayValue(workHours.getHOUR_14()));
                dataRow.createCell(17).setCellValue(convertToDisplayValue(workHours.getHOUR_15()));
                dataRow.createCell(18).setCellValue(convertToDisplayValue(workHours.getHOUR_16()));
                dataRow.createCell(19).setCellValue(convertToDisplayValue(workHours.getHOUR_17()));
                dataRow.createCell(20).setCellValue(convertToDisplayValue(workHours.getHOUR_18()));
                dataRow.createCell(21).setCellValue(convertToDisplayValue(workHours.getHOUR_19()));
                dataRow.createCell(22).setCellValue(convertToDisplayValue(workHours.getHOUR_20()));
                dataRow.createCell(23).setCellValue(convertToDisplayValue(workHours.getHOUR_21()));
                dataRow.createCell(24).setCellValue(convertToDisplayValue(workHours.getHOUR_22()));
                dataRow.createCell(25).setCellValue(convertToDisplayValue(workHours.getHOUR_23()));
                dataRow.createCell(26).setCellValue(convertToDisplayValue(workHours.getHOUR_24()));

                // DESCRIPTION column
                Cell descriptionCell = dataRow.createCell(27);
                descriptionCell.setCellValue(workHours.getDESCRIPTION() != null ? workHours.getDESCRIPTION().toString() : "");
                descriptionCell.setCellStyle(borderStyle);

                // Apply border style for each cell in the row
                for (int j = 0; j < columns.size(); j++) {
                    dataRow.getCell(j).setCellStyle(borderStyle);
                }
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export WorkDayHours data.");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }



    private String convertToDisplayValue(BigDecimal value) {
        return (value != null && value.compareTo(BigDecimal.ONE) == 0) ? "v" : "";
    }
    
 // Update hours for SHIFT 1 (HOUR_1 to HOUR_8)
    public Optional<DWorkDayHours> updateHoursShift3(BigDecimal newValue, Date parsedDate, String description) {
        Optional<DWorkDayHours> workHoursOpt = getWorkDayHoursByDateDesc(parsedDate, description);
        if (workHoursOpt.isPresent()) {
            DWorkDayHours workHours = workHoursOpt.get();

            workHours.setHOUR_1(newValue);
            workHours.setHOUR_2(newValue);
            workHours.setHOUR_3(newValue);
            workHours.setHOUR_4(newValue);
            workHours.setHOUR_5(newValue);
            workHours.setHOUR_6(newValue);
            workHours.setHOUR_7(newValue);
            workHours.setHOUR_8(newValue);

            workHours.setLAST_UPDATE_DATE(new Date());
            dWorkDayHoursRepo.save(workHours);

            System.out.println("Shift 1 hours updated successfully.");
            return Optional.of(workHours);
        } else {
            System.out.println("No WorkDayHours record found for date: " + parsedDate);
            return Optional.empty();
        }
    }

    // Update hours for SHIFT 1 (HOUR_9 to HOUR_16)
    public Optional<DWorkDayHours> updateHoursShift1(BigDecimal newValue, Date parsedDate, String description) {
        Optional<DWorkDayHours> workHoursOpt = getWorkDayHoursByDateDesc(parsedDate, description);
        if (workHoursOpt.isPresent()) {
            DWorkDayHours workHours = workHoursOpt.get();

            workHours.setHOUR_9(newValue);
            workHours.setHOUR_10(newValue);
            workHours.setHOUR_11(newValue);
            workHours.setHOUR_12(newValue);
            workHours.setHOUR_13(newValue);
            workHours.setHOUR_14(newValue);
            workHours.setHOUR_15(newValue);
            workHours.setHOUR_16(newValue);

            workHours.setLAST_UPDATE_DATE(new Date());
            dWorkDayHoursRepo.save(workHours);

            System.out.println("Shift 2 hours updated successfully.");
            return Optional.of(workHours);
        } else {
            System.out.println("No WorkDayHours record found for date: " + parsedDate);
            return Optional.empty();
        }
    }

    // Update hours for SHIFT 3 (HOUR_17 to HOUR_24)
    public Optional<DWorkDayHours> updateHoursShift2(BigDecimal newValue, Date parsedDate, String description) {
        Optional<DWorkDayHours> workHoursOpt = getWorkDayHoursByDateDesc(parsedDate, description);
        if (workHoursOpt.isPresent()) {
            DWorkDayHours workHours = workHoursOpt.get();

            workHours.setHOUR_17(newValue);
            workHours.setHOUR_18(newValue);
            workHours.setHOUR_19(newValue);
            workHours.setHOUR_20(newValue);
            workHours.setHOUR_21(newValue);
            workHours.setHOUR_22(newValue);
            workHours.setHOUR_23(newValue);
            workHours.setHOUR_24(newValue);

            workHours.setLAST_UPDATE_DATE(new Date());
            dWorkDayHoursRepo.save(workHours);

            System.out.println("Shift 3 hours updated successfully.");
            return Optional.of(workHours);
        } else {
            System.out.println("No WorkDayHours record found for date: " + parsedDate);
            return Optional.empty();
        }
    }
    
    public DWorkDayHours turnOffHour(String dateWd, String hour, String description) {
        return updateHourValue(dateWd, hour, BigDecimal.ZERO, description);
    }

    public DWorkDayHours turnOnHour(String dateWd, String hour, String description) {
        return updateHourValue(dateWd, hour, BigDecimal.ONE, description);
    }

    public DWorkDayHours updateHourValue(String dateWd, String hour, BigDecimal newValue, String description) {
        try {
            Optional<DWorkDayHours> currentWorkHoursOpt = dWorkDayHoursRepo.findDWdHoursByDateAndDescription(dateWd, description);

            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
            LocalDate parsedDate = LocalDate.parse(dateWd, formatter);

            currentWorkHoursOpt.ifPresentOrElse(currentDWorkDayHours -> {
                setHourValue(currentDWorkDayHours, hour, newValue);
                updateOffAndSemiOff(currentDWorkDayHours, dateWd);

                dWorkDayHoursRepo.save(currentDWorkDayHours);
            }, () -> {
                throw new RuntimeException("WorkDay with date " + dateWd + " not found.");
            });

            return currentWorkHoursOpt.orElseThrow(() -> new RuntimeException("WorkDay with date " + dateWd + " not found."));
        } catch (DateTimeParseException e) {
            System.err.println("Error parsing date: " + e.getMessage());
            throw new RuntimeException("Invalid date format provided: " + dateWd, e);
        } catch (Exception e) {
            System.err.println("Error updating hour: " + e.getMessage());
            throw e;
        }
    }

    private void setHourValue(DWorkDayHours dWorkDayHours, String hour, BigDecimal newValue) {
        switch (hour.toUpperCase()) {
            case "HOUR_1": dWorkDayHours.setHOUR_1(newValue); break;
            case "HOUR_2": dWorkDayHours.setHOUR_2(newValue); break;
            case "HOUR_3": dWorkDayHours.setHOUR_3(newValue); break;
            case "HOUR_4": dWorkDayHours.setHOUR_4(newValue); break;
            case "HOUR_5": dWorkDayHours.setHOUR_5(newValue); break;
            case "HOUR_6": dWorkDayHours.setHOUR_6(newValue); break;
            case "HOUR_7": dWorkDayHours.setHOUR_7(newValue); break;
            case "HOUR_8": dWorkDayHours.setHOUR_8(newValue); break;
            case "HOUR_9": dWorkDayHours.setHOUR_9(newValue); break;
            case "HOUR_10": dWorkDayHours.setHOUR_10(newValue); break;
            case "HOUR_11": dWorkDayHours.setHOUR_11(newValue); break;
            case "HOUR_12": dWorkDayHours.setHOUR_12(newValue); break;
            case "HOUR_13": dWorkDayHours.setHOUR_13(newValue); break;
            case "HOUR_14": dWorkDayHours.setHOUR_14(newValue); break;
            case "HOUR_15": dWorkDayHours.setHOUR_15(newValue); break;
            case "HOUR_16": dWorkDayHours.setHOUR_16(newValue); break;
            case "HOUR_17": dWorkDayHours.setHOUR_17(newValue); break;
            case "HOUR_18": dWorkDayHours.setHOUR_18(newValue); break;
            case "HOUR_19": dWorkDayHours.setHOUR_19(newValue); break;
            case "HOUR_20": dWorkDayHours.setHOUR_20(newValue); break;
            case "HOUR_21": dWorkDayHours.setHOUR_21(newValue); break;
            case "HOUR_22": dWorkDayHours.setHOUR_22(newValue); break;
            case "HOUR_23": dWorkDayHours.setHOUR_23(newValue); break;
            case "HOUR_24": dWorkDayHours.setHOUR_24(newValue); break;
            default: throw new IllegalArgumentException("Invalid hour specified: " + hour);
        }
    }
    private void updateOffAndSemiOff(DWorkDayHours dWorkDayHours, String dateWd) {
        Optional<WorkDay> workDayOpt = workDayRepo.findByDDateWd(dateWd);
        workDayOpt.ifPresentOrElse(workDay -> {
            BigDecimal shift1 = Stream.of(
                    dWorkDayHours.getHOUR_9(), dWorkDayHours.getHOUR_10(), dWorkDayHours.getHOUR_11(), dWorkDayHours.getHOUR_12(),
                    dWorkDayHours.getHOUR_13(), dWorkDayHours.getHOUR_14(), dWorkDayHours.getHOUR_15(), dWorkDayHours.getHOUR_16()
                ).reduce(BigDecimal.ZERO, BigDecimal::add);

            BigDecimal shift2 = Stream.of(
                    dWorkDayHours.getHOUR_17(), dWorkDayHours.getHOUR_18(), dWorkDayHours.getHOUR_19(),
                    dWorkDayHours.getHOUR_20(), dWorkDayHours.getHOUR_21(), dWorkDayHours.getHOUR_22(),
                    dWorkDayHours.getHOUR_23(), dWorkDayHours.getHOUR_24()
                ).reduce(BigDecimal.ZERO, BigDecimal::add);

            BigDecimal shift3 = Stream.of(
                    dWorkDayHours.getHOUR_1(), dWorkDayHours.getHOUR_2(), dWorkDayHours.getHOUR_3(),
                    dWorkDayHours.getHOUR_4(), dWorkDayHours.getHOUR_5(), dWorkDayHours.getHOUR_6(),
                    dWorkDayHours.getHOUR_7(), dWorkDayHours.getHOUR_8()
                ).reduce(BigDecimal.ZERO, BigDecimal::add);

            // Log nilai shift
            System.out.println("Shift1: " + shift1);
            System.out.println("Shift2: " + shift2);
            System.out.println("Shift3: " + shift3);

            DayOfWeek dayOfWeek = getDayOfWeekFromDate(dateWd);

            // Jika hari Sabtu atau Minggu, set OFF
            if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
                workDay.setOFF(BigDecimal.ONE);
                workDay.setSEMI_OFF(BigDecimal.ZERO);
                
            } else if (dayOfWeek == DayOfWeek.MONDAY) {
                if (shift1.compareTo(BigDecimal.valueOf(8)) == 0 && shift2.compareTo(BigDecimal.valueOf(8)) == 0) {
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
                } else if (shift1.compareTo(BigDecimal.valueOf(8)) == 0 && shift2.compareTo(BigDecimal.valueOf(8)) == 0 && shift3.compareTo(BigDecimal.valueOf(8)) > 0) {
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
