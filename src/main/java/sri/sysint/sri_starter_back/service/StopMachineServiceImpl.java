package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.StopMachine;
import sri.sysint.sri_starter_back.repository.StopMachineRepo;

@Service
@Transactional
public class StopMachineServiceImpl {

    @Autowired
    private StopMachineRepo stopMachineRepo;

    public StopMachineServiceImpl(StopMachineRepo stopMachineRepo) {
        this.stopMachineRepo = stopMachineRepo;
    }

    public BigDecimal getNewId() {
        return stopMachineRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<StopMachine> getAllStopMachines() {
        Iterable<StopMachine> stopMachines = stopMachineRepo.getDataOrderId();
        List<StopMachine> stopMachineList = new ArrayList<>();
        for (StopMachine item : stopMachines) {
            stopMachineList.add(new StopMachine(item));
        }
        return stopMachineList;
    }

    public Optional<StopMachine> getStopMachineById(BigDecimal id) {
        return stopMachineRepo.findById(id);
    }

    public StopMachine saveStopMachine(StopMachine stopMachine) {
        try {
            LocalDateTime startDate = stopMachine.getDATE_PM().toInstant()
                    .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
            
            stopMachine.setDATE_PM(Date.from(startDate.atZone(ZoneId.of("UTC")).toInstant()));
            stopMachine.setTOTAL_TIME(calculateTotalTime(stopMachine.getSTART_TIME(), stopMachine.getEND_TIME()));
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
                
                LocalDateTime startDate = stopMachine.getDATE_PM().toInstant()
                        .atZone(ZoneId.of("UTC")).toLocalDateTime().withHour(17).withMinute(0).withSecond(0);
                
                currentStopMachine.setSTART_TIME(stopMachine.getSTART_TIME());
                currentStopMachine.setEND_TIME(stopMachine.getEND_TIME());;

                currentStopMachine.setTOTAL_TIME(calculateTotalTime(stopMachine.getSTART_TIME(), stopMachine.getEND_TIME()));
                
                currentStopMachine.setDATE_PM(Date.from(startDate.atZone(ZoneId.of("UTC")).toInstant()));
                
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
            "DATE_PM",
            "START_TIME",
            "END_TIME",
            "TOTAL_TIME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("STOP MACHINE DATA");

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

            // Time style for formatting time (hh:mm)
            CellStyle timeStyle = workbook.createCellStyle();
            timeStyle.cloneStyleFrom(borderStyle);
            timeStyle.setDataFormat(dateFormat.getFormat("hh:mm"));

            // Set column widths (adjusted from your example)
            sheet.setColumnWidth(0, 10 * 256); 
            sheet.setColumnWidth(1, 20 * 256); 
            sheet.setColumnWidth(2, 30 * 256); 
            sheet.setColumnWidth(3, 20 * 256); 
            sheet.setColumnWidth(4, 20 * 256); 
            sheet.setColumnWidth(5, 20 * 256); 
            sheet.setColumnWidth(6, 20 * 256);

            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            // Fill data rows
            int rowIndex = 1;
            int nomor = 1;
            for (StopMachine sm : stopMachines) {
                Row dataRow = sheet.createRow(rowIndex++);

                // NOMOR column
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                // STOP_MACHINE_ID column
                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(sm.getSTOP_MACHINE_ID() != null ? sm.getSTOP_MACHINE_ID().doubleValue() : null);  // Handle null
                idCell.setCellStyle(borderStyle);

                // WORK_CENTER_TEXT column
                Cell workCenterCell = dataRow.createCell(2);
                workCenterCell.setCellValue(sm.getWORK_CENTER_TEXT() != null ? sm.getWORK_CENTER_TEXT() : "");  // Handle null
                workCenterCell.setCellStyle(borderStyle);

                // DATE_PM column (formatted as date)
                Cell datePmCell = dataRow.createCell(3);
                if (sm.getDATE_PM() != null) {
                    datePmCell.setCellValue(sm.getDATE_PM());
                    datePmCell.setCellStyle(dateStyle);
                } else {
                    datePmCell.setCellValue("");
                    datePmCell.setCellStyle(borderStyle);
                }

                // START_TIME column (formatted as time)
                Cell startTimeCell = dataRow.createCell(4);
                if (sm.getSTART_TIME() != null) {
                    startTimeCell.setCellValue(sm.getSTART_TIME());
                    startTimeCell.setCellStyle(timeStyle);
                } else {
                    startTimeCell.setCellValue("");
                    startTimeCell.setCellStyle(borderStyle);
                }

                // END_TIME column (formatted as time)
                Cell endTimeCell = dataRow.createCell(5);
                if (sm.getEND_TIME() != null) {
                    endTimeCell.setCellValue(sm.getEND_TIME());
                    endTimeCell.setCellStyle(timeStyle);
                } else {
                    endTimeCell.setCellValue("");
                    endTimeCell.setCellStyle(borderStyle);
                }

                // TOTAL_TIME column (calculated as difference between END_TIME and START_TIME)
                Cell totalTimeCell = dataRow.createCell(6);
                if (sm.getSTART_TIME() != null && sm.getEND_TIME() != null) {
                    totalTimeCell.setCellValue(sm.getTOTAL_TIME().doubleValue());
                } else {
                    totalTimeCell.setCellValue("");
                }
                totalTimeCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export data");
            return null;
        } finally {
            workbook.close();
            out.close();
        }
    }

    
    private int combineTimeWithMinutes(String timeStr) {
        // Mengonversi string waktu "HH:mm" menjadi total menit
        String[] parts = timeStr.split(":");
        int hours = Integer.parseInt(parts[0]);
        int minutes = Integer.parseInt(parts[1]);
        return hours * 60 + minutes;
    }
    
    private BigDecimal calculateTotalTime(String startTimeStr, String endTimeStr) {
        int startMinutes = combineTimeWithMinutes(startTimeStr);
        int endMinutes = combineTimeWithMinutes(endTimeStr);
        
        int diffMinutes = endMinutes - startMinutes;
        
        if (diffMinutes < 0) {
            diffMinutes += 24 * 60; 
        }

        return BigDecimal.valueOf(diffMinutes);
    }

}
