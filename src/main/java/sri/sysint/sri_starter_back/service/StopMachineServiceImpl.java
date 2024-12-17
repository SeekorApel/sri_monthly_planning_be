package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.Duration;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
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

import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineTass;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.model.StopMachine;
import sri.sysint.sri_starter_back.repository.StopMachineRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;
import sri.sysint.sri_starter_back.repository.MachineTassRepo;

@Service
@Transactional
public class StopMachineServiceImpl {

    @Autowired
    private StopMachineRepo stopMachineRepo;
    @Autowired
    private MachineCuringRepo machineCuringRepo;
    @Autowired
    private MachineTassRepo machineTassRepo;
    public StopMachineServiceImpl(StopMachineRepo stopMachineRepo) {
        this.stopMachineRepo = stopMachineRepo;
    }

    public BigDecimal getNewId() {
        return stopMachineRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<StopMachine> getAllStopMachines() {        
        Iterable<StopMachine> stopMachines = stopMachineRepo.findAll();
        List<StopMachine> stopMachineList = new ArrayList<>();
//        for (StopMachine item : stopMachines) {
//        	StopMachine stopMachineTemp = new StopMachine(item);
//        	stopMachineList.add(stopMachineTemp);
//            System.out.println(stopMachineTemp.getEND_TIME());
//            System.out.println(stopMachineTemp.getSTART_TIME());
//        }
        
        for (StopMachine item : stopMachines) {
            System.out.println("Original Start Time: " + item.getSTART_TIME());
            StopMachine stopMachineTemp = new StopMachine(item);
            System.out.println("After Copy Start Time: " + stopMachineTemp.getSTART_TIME());
            stopMachineList.add(stopMachineTemp);
        }

        System.out.println(stopMachineList);
        return stopMachineList;
    }

    public Optional<StopMachine> getStopMachineById(BigDecimal id) {
        return stopMachineRepo.findById(id);
    }

    public StopMachine saveStopMachine(StopMachine stopMachine) {
        try {
            LocalDateTime startDate = stopMachine.getSTART_DATE().toInstant()
                    .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
            LocalDateTime endDate = stopMachine.getEND_DATE().toInstant()
                    .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
            
            stopMachine.setSTART_DATE(Date.from(startDate.atZone(ZoneId.of("UTC")).toInstant()));
            stopMachine.setEND_DATE(Date.from(endDate.atZone(ZoneId.of("UTC")).toInstant()));
            stopMachine.setTOTAL_TIME(calculateTotalTime(stopMachine.getSTART_TIME(), stopMachine.getEND_TIME(), stopMachine.getSTART_DATE().toString(), stopMachine.getEND_DATE().toString()));
            stopMachine.setSTOP_MACHINE_ID(getNewId());
            stopMachine.setSTATUS(BigDecimal.valueOf(1));
            stopMachine.setCREATION_DATE(new Date());
            stopMachine.setLAST_UPDATE_DATE(new Date());
            return stopMachineRepo.save(stopMachine);
        } catch (Exception e) {
            System.err.println("Error saving StopMachine: " + e.getMessage());
            throw e;
        }
    }

    public StopMachine updateStopMachine(StopMachine stopMachine) {
        try {
            Optional<StopMachine> currentStopMachineOpt = stopMachineRepo.findById(stopMachine.getSTOP_MACHINE_ID());

            if (currentStopMachineOpt.isPresent()) {
                StopMachine currentStopMachine = currentStopMachineOpt.get();
                
                LocalDateTime startDate = stopMachine.getSTART_DATE().toInstant()
                        .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
                LocalDateTime endDate = stopMachine.getEND_DATE().toInstant()
                        .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
                
                currentStopMachine.setSTART_TIME(stopMachine.getSTART_TIME());
                currentStopMachine.setEND_TIME(stopMachine.getEND_TIME());;

                currentStopMachine.setTOTAL_TIME(calculateTotalTime(stopMachine.getSTART_TIME(), stopMachine.getEND_TIME(), stopMachine.getSTART_DATE().toString(), stopMachine.getEND_DATE().toString()));
                
                currentStopMachine.setSTART_DATE(Date.from(startDate.atZone(ZoneId.of("UTC")).toInstant()));
                currentStopMachine.setEND_DATE(Date.from(endDate.atZone(ZoneId.of("UTC")).toInstant()));

                currentStopMachine.setWORK_CENTER_TEXT(stopMachine.getWORK_CENTER_TEXT());
                currentStopMachine.setLAST_UPDATE_DATE(new Date());
                currentStopMachine.setLAST_UPDATED_BY(stopMachine.getLAST_UPDATED_BY());

                return stopMachineRepo.save(currentStopMachine);
            } else {
                throw new RuntimeException("StopMachine with ID " + stopMachine.getSTOP_MACHINE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating StopMachine: " + e.getMessage());
            throw e;
        }
    }

    public StopMachine deleteStopMachine(StopMachine stopMachine) {
        try {
            Optional<StopMachine> currentStopMachineOpt = stopMachineRepo.findById(stopMachine.getSTOP_MACHINE_ID());

            if (currentStopMachineOpt.isPresent()) {
                StopMachine currentStopMachine = currentStopMachineOpt.get();
                currentStopMachine.setSTATUS(BigDecimal.valueOf(0));
                currentStopMachine.setLAST_UPDATE_DATE(new Date());
                currentStopMachine.setLAST_UPDATED_BY(stopMachine.getLAST_UPDATED_BY());

                return stopMachineRepo.save(currentStopMachine);
            } else {
                throw new RuntimeException("StopMachine with ID " + stopMachine.getSTOP_MACHINE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error deleting StopMachine: " + e.getMessage());
            throw e;
        }
    }

    public StopMachine restoreStopMachine(StopMachine stopMachine) {
        try {
            Optional<StopMachine> currentStopMachineOpt = stopMachineRepo.findById(stopMachine.getSTOP_MACHINE_ID());

            if (currentStopMachineOpt.isPresent()) {
                StopMachine currentStopMachine = currentStopMachineOpt.get();
                currentStopMachine.setSTATUS(BigDecimal.valueOf(1));
                currentStopMachine.setLAST_UPDATE_DATE(new Date());
                currentStopMachine.setLAST_UPDATED_BY(stopMachine.getLAST_UPDATED_BY());

                return stopMachineRepo.save(currentStopMachine);
            } else {
                throw new RuntimeException("StopMachine with ID " + stopMachine.getSTOP_MACHINE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring StopMachine: " + e.getMessage());
            throw e;
        }
    }

    public List<StopMachine> getActiveStopMachines() {
        return stopMachineRepo.findStopMachineActive();
    }

    public void deleteAllStopMachines() {
        stopMachineRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportStopMachinesExcel() throws IOException {
        List<StopMachine> stopMachines = stopMachineRepo.getDataOrderId();
        return dataToExcel(stopMachines);
    }

    public ByteArrayInputStream dataToExcel(List<StopMachine> stopMachines) throws IOException {
        String[] header = {
            "NOMOR", 
            "STOP_MACHINE_ID", 
            "WORK_CENTER_TEXT", 
            "START_DATE",
            "END_DATE",
            "START_TIME",
            "END_TIME",
            "TOTAL_TIME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<String> machineCuringNames = machineCuringRepo.findMachineCuringActive()
                .stream()
                .map(MachineCuring::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());

            List<String> machineTassNames = machineTassRepo.findMachineTassActive()
                .stream()
                .map(MachineTass::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());

            List<String> combinedMachineNames = new ArrayList<>();
            combinedMachineNames.addAll(machineCuringNames);
            combinedMachineNames.addAll(machineTassNames);

            Sheet sheet = workbook.createSheet("STOP MACHINE DATA");

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

            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.cloneStyleFrom(borderStyle);
            DataFormat dateFormat = workbook.createDataFormat();
            dateStyle.setDataFormat(dateFormat.getFormat("dd-MM-yyyy"));

            CellStyle timeStyle = workbook.createCellStyle();
            timeStyle.cloneStyleFrom(borderStyle);
            timeStyle.setDataFormat(dateFormat.getFormat("hh:mm"));

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_MACHINES");
            for (int i = 0; i < combinedMachineNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(combinedMachineNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("MachineNames");
            namedRange.setRefersToFormula("HIDDEN_MACHINES!$A$1:$A$" + combinedMachineNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            int nomor = 1;
            for (StopMachine sm : stopMachines) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(sm.getSTOP_MACHINE_ID() != null ? sm.getSTOP_MACHINE_ID().doubleValue() : null);
                idCell.setCellStyle(borderStyle);

                Cell workCenterCell = dataRow.createCell(2);
                workCenterCell.setCellValue(sm.getWORK_CENTER_TEXT() != null ? sm.getWORK_CENTER_TEXT() : "");
                workCenterCell.setCellStyle(borderStyle);

                Cell startdatePmCell = dataRow.createCell(3);
                if (sm.getSTART_DATE() != null) {
                    startdatePmCell.setCellValue(sm.getSTART_DATE());
                    startdatePmCell.setCellStyle(dateStyle);
                } else {
                    startdatePmCell.setCellValue("");
                    startdatePmCell.setCellStyle(borderStyle);
                }

                Cell enddatePmCell = dataRow.createCell(4);
                if (sm.getEND_DATE() != null) {
                    enddatePmCell.setCellValue(sm.getEND_DATE());
                    enddatePmCell.setCellStyle(dateStyle);
                } else {
                    enddatePmCell.setCellValue("");
                    enddatePmCell.setCellStyle(borderStyle);
                }

                Cell startTimeCell = dataRow.createCell(5);
                if (sm.getSTART_TIME() != null) {
                    startTimeCell.setCellValue(sm.getSTART_TIME());
                    startTimeCell.setCellStyle(timeStyle);
                } else {
                    startTimeCell.setCellValue("");
                    startTimeCell.setCellStyle(borderStyle);
                }

                Cell endTimeCell = dataRow.createCell(6);
                if (sm.getEND_TIME() != null) {
                    endTimeCell.setCellValue(sm.getEND_TIME());
                    endTimeCell.setCellStyle(timeStyle);
                } else {
                    endTimeCell.setCellValue("");
                    endTimeCell.setCellStyle(borderStyle);
                }

                Cell totalTimeCell = dataRow.createCell(7);
                if (sm.getSTART_TIME() != null && sm.getEND_TIME() != null) {
                    totalTimeCell.setCellValue(calculateTotalTimeExcel(sm.getSTART_TIME(), sm.getEND_TIME(), sm.getSTART_DATE().toString(), sm.getEND_DATE().toString()).doubleValue());
                } else {
                    totalTimeCell.setCellValue("");
                }
                totalTimeCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("MachineNames");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    
    private int combineTimeWithMinutes(String timeStr) {
        String[] parts = timeStr.split(":");
        int hours = Integer.parseInt(parts[0]);
        int minutes = Integer.parseInt(parts[1]);
        return hours * 60 + minutes;
    }


    // Method untuk menghitung total waktu
    public BigDecimal calculateTotalTime(String startTimeStr, String endTimeStr, String startDateStr, String endDateStr) {
        try {
            System.out.println("startTimeStr: " + startTimeStr);
            System.out.println("endTimeStr: " + endTimeStr);
            System.out.println("startDateStr: " + startDateStr);
            System.out.println("endDateStr: " + endDateStr);

            // Formatter untuk format input tanggal
            DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("EEE MMM dd HH:mm:ss z yyyy");
            
            // Parsing string tanggal menggunakan formatter
            ZonedDateTime startDateTimeParsed = ZonedDateTime.parse(startDateStr, inputFormatter);
            ZonedDateTime endDateTimeParsed = ZonedDateTime.parse(endDateStr, inputFormatter);

            // Ubah ke LocalDate
            LocalDate startDate = startDateTimeParsed.toLocalDate();
            LocalDate endDate = endDateTimeParsed.toLocalDate();

            // Parsing waktu
            LocalTime startTime = LocalTime.parse(startTimeStr); // Format: HH:mm
            LocalTime endTime = LocalTime.parse(endTimeStr);

            // Gabungkan tanggal dan waktu
            LocalDateTime startDateTime = LocalDateTime.of(startDate, startTime);
            LocalDateTime endDateTime = LocalDateTime.of(endDate, endTime);

            // Hitung selisih waktu dalam menit
            long totalMinutes = Duration.between(startDateTime, endDateTime).toMinutes();

            // Jika totalMinutes negatif, artinya tanggal/waktu di-input salah
            if (totalMinutes < 0) {
                throw new IllegalArgumentException("End date and time must be after start date and time.");
            }

            System.out.println("Total Time in Minutes: " + totalMinutes);
            return BigDecimal.valueOf(totalMinutes);

        } catch (DateTimeParseException e) {
            System.err.println("Error parsing date or time: " + e.getMessage());
            return BigDecimal.ZERO;
        } catch (Exception e) {
            System.err.println("Error calculating total time: " + e.getMessage());
            return BigDecimal.ZERO;
        }
    }
    
    public BigDecimal calculateTotalTimeExcel(String startTimeStr, String endTimeStr, String startDateStr, String endDateStr) {
        try {
            System.out.println("startTimeStr: " + startTimeStr);
            System.out.println("endTimeStr: " + endTimeStr);
            System.out.println("startDateStr: " + startDateStr);
            System.out.println("endDateStr: " + endDateStr);

            // Formatter untuk format input tanggal (dengan milidetik .S)
            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.S");
            
            // Parsing string tanggal menggunakan formatter
            LocalDateTime startDateTimeParsed = LocalDateTime.parse(startDateStr, dateFormatter);
            LocalDateTime endDateTimeParsed = LocalDateTime.parse(endDateStr, dateFormatter);

            // Ubah waktu sesuai startTimeStr dan endTimeStr
            LocalDate startDate = startDateTimeParsed.toLocalDate();
            LocalDate endDate = endDateTimeParsed.toLocalDate();

            LocalTime startTime = LocalTime.parse(startTimeStr); // Format: HH:mm
            LocalTime endTime = LocalTime.parse(endTimeStr);

            // Gabungkan tanggal dan waktu
            LocalDateTime startDateTime = LocalDateTime.of(startDate, startTime);
            LocalDateTime endDateTime = LocalDateTime.of(endDate, endTime);

            // Hitung selisih waktu dalam menit
            long totalMinutes = Duration.between(startDateTime, endDateTime).toMinutes();

            // Jika totalMinutes negatif, artinya tanggal/waktu di-input salah
            if (totalMinutes < 0) {
                throw new IllegalArgumentException("End date and time must be after start date and time.");
            }

            System.out.println("Total Time in Minutes: " + totalMinutes);
            return BigDecimal.valueOf(totalMinutes);

        } catch (DateTimeParseException e) {
            System.err.println("Error parsing date or time: " + e.getMessage());
            return BigDecimal.ZERO;
        } catch (Exception e) {
            System.err.println("Error calculating total time: " + e.getMessage());
            return BigDecimal.ZERO;
        }
    }
    
    public ByteArrayInputStream layoutStopMachinesExcel() throws IOException {
        return layoutToExcel();
    }

    public ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR", 
            "STOP_MACHINE_ID", 
            "WORK_CENTER_TEXT", 
            "START_DATE",
            "END_DATE",
            "START_TIME",
            "END_TIME",
            "TOTAL_TIME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<String> machineCuringNames = machineCuringRepo.findMachineCuringActive()
                .stream()
                .map(MachineCuring::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());

            List<String> machineTassNames = machineTassRepo.findMachineTassActive()
                .stream()
                .map(MachineTass::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());

            List<String> combinedMachineNames = new ArrayList<>();
            combinedMachineNames.addAll(machineCuringNames);
            combinedMachineNames.addAll(machineTassNames);

            Sheet sheet = workbook.createSheet("STOP MACHINE DATA");

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

            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.cloneStyleFrom(borderStyle);
            DataFormat dateFormat = workbook.createDataFormat();
            dateStyle.setDataFormat(dateFormat.getFormat("dd-MM-yyyy"));

            CellStyle timeStyle = workbook.createCellStyle();
            timeStyle.cloneStyleFrom(borderStyle);
            timeStyle.setDataFormat(dateFormat.getFormat("hh:mm"));

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_MACHINES");
            for (int i = 0; i < combinedMachineNames.size(); i++) {
                Row row = hiddenSheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(combinedMachineNames.get(i));
            }

            Name namedRange = workbook.createName();
            namedRange.setNameName("MachineNames");
            namedRange.setRefersToFormula("HIDDEN_MACHINES!$A$1:$A$" + combinedMachineNames.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            int nomor = 1;
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }
            
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("MachineNames");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
}
