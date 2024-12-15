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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.MachineAllowence;
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineTass;
import sri.sysint.sri_starter_back.repository.MachineAllowenceRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;
import sri.sysint.sri_starter_back.repository.MachineTassRepo;


@Service
@Transactional
public class MachineAllowenceServiceImpl {
	@Autowired
    private MachineAllowenceRepo machineAllowenceRepo;

    @Autowired
    private MachineCuringRepo machineCuringRepo;

    @Autowired
    private MachineTassRepo machineTassRepo;
	
    public MachineAllowenceServiceImpl(MachineAllowenceRepo machineAllowenceRepo){
        this.machineAllowenceRepo = machineAllowenceRepo;
    }
    
    public BigDecimal getNewId() {
    	return machineAllowenceRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<MachineAllowence> getAllMachineAllowence() {
    	Iterable<MachineAllowence> machineAllowences = machineAllowenceRepo.getDataOrderId();
        List<MachineAllowence> machineAllowenceList = new ArrayList<>();
        for (MachineAllowence item : machineAllowences) {
            MachineAllowence machineAllowenceTemp = new MachineAllowence(item);
            machineAllowenceList.add(machineAllowenceTemp);
        }
        
        return machineAllowenceList;
    }
    
    public Optional<MachineAllowence> getMachineAllowenceById(BigDecimal id) {
    	Optional<MachineAllowence> machineAllowence = machineAllowenceRepo.findById(id);
    	return machineAllowence;
    }
    
    public MachineAllowence saveMachineAllowence(MachineAllowence machineAllowence) {
        try {
        	machineAllowence.setMACHINE_ALLOW_ID(getNewId());
            machineAllowence.setSTATUS(BigDecimal.valueOf(1));
            machineAllowence.setCREATION_DATE(new Date());
            machineAllowence.setLAST_UPDATE_DATE(new Date());
            return machineAllowenceRepo.save(machineAllowence);
        } catch (Exception e) {
            System.err.println("Error saving machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineAllowence updateMachineAllowence(MachineAllowence machineAllowence) {
        try {
            Optional<MachineAllowence> currentMachineAllowenceOpt = machineAllowenceRepo.findById(machineAllowence.getMACHINE_ALLOW_ID());
            
            if (currentMachineAllowenceOpt.isPresent()) {
                MachineAllowence currentMachineAllowence = currentMachineAllowenceOpt.get();
                
                currentMachineAllowence.setID_MACHINE(machineAllowence.getID_MACHINE());
                currentMachineAllowence.setPERSON_RESPONSIBLE(machineAllowence.getPERSON_RESPONSIBLE());
                currentMachineAllowence.setSHIFT_1(machineAllowence.getSHIFT_1());
                currentMachineAllowence.setSHIFT_2(machineAllowence.getSHIFT_2());
                currentMachineAllowence.setSHIFT_3(machineAllowence.getSHIFT_3());
                currentMachineAllowence.setSHIFT_1_FRIDAY(machineAllowence.getSHIFT_1_FRIDAY());
                currentMachineAllowence.setTOTAL_SHIFT_123(machineAllowence.getTOTAL_SHIFT_123());
                currentMachineAllowence.setLAST_UPDATE_DATE(new Date());
                currentMachineAllowence.setLAST_UPDATED_BY(machineAllowence.getLAST_UPDATED_BY());
                
                return machineAllowenceRepo.save(currentMachineAllowence);
            } else {
                throw new RuntimeException("MachineAllowence with ID " + machineAllowence.getMACHINE_ALLOW_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineAllowence deleteMachineAllowence(MachineAllowence machineAllowence) {
        try {
            Optional<MachineAllowence> currentMachineAllowenceOpt = machineAllowenceRepo.findById(machineAllowence.getMACHINE_ALLOW_ID());
            
            if (currentMachineAllowenceOpt.isPresent()) {
                MachineAllowence currentMachineAllowence = currentMachineAllowenceOpt.get();
                
                currentMachineAllowence.setSTATUS(BigDecimal.valueOf(0));
                currentMachineAllowence.setLAST_UPDATE_DATE(new Date());
                currentMachineAllowence.setLAST_UPDATED_BY(machineAllowence.getLAST_UPDATED_BY());
                
                return machineAllowenceRepo.save(currentMachineAllowence);
            } else {
                throw new RuntimeException("MachineAllowence with ID " + machineAllowence.getMACHINE_ALLOW_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllMachineAllowence() {
    	machineAllowenceRepo.deleteAll();
    }
    
    public MachineAllowence restoreMachineAllowence(MachineAllowence machineAllowence) {
        try {
            Optional<MachineAllowence> currentMachineAllowenceOpt = machineAllowenceRepo.findById(machineAllowence.getMACHINE_ALLOW_ID());
            
            if (currentMachineAllowenceOpt.isPresent()) {
                MachineAllowence currentMachineAllowence = currentMachineAllowenceOpt.get();
                
                currentMachineAllowence.setSTATUS(BigDecimal.valueOf(1));
                currentMachineAllowence.setLAST_UPDATE_DATE(new Date());
                currentMachineAllowence.setLAST_UPDATED_BY(machineAllowence.getLAST_UPDATED_BY());
                
                return machineAllowenceRepo.save(currentMachineAllowence);
            } else {
                throw new RuntimeException("MachineAllowence with ID " + machineAllowence.getMACHINE_ALLOW_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineAllowence: " + e.getMessage());
            throw e;
        }
    }
    
    public ByteArrayInputStream exportMachineAllowenceExcel() throws IOException {
        List<MachineAllowence> machineAllowences = getAllMachineAllowence(); // Mengambil semua data MachineAllowence
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineAllowences); // Konversi data ke Excel
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineAllowence> machineAllowences) throws IOException {
        String[] header = {
            "NOMOR",
            "MACHINE_ALLOW_ID",
            "ID_MACHINE",
            "PERSON_RESPONSIBLE",
            "SHIFT_1",
            "SHIFT_2",
            "SHIFT_3",
            "SHIFT_1_FRIDAY",
            "TOTAL_SHIFT_123"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<MachineCuring> activeMachineCurings = machineCuringRepo.findMachineCuringActive();
            List<MachineTass> activeMachineTasss = machineTassRepo.findMachineTassActive();
            List<String> machineCuringWCT = activeMachineCurings.stream()
                .map(MachineCuring::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());
            List<String> machineTassWCT = activeMachineTasss.stream()
                .map(MachineTass::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());

            Sheet sheet = workbook.createSheet("MACHINE ALLOWENCE DATA");
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

            Sheet hiddenMachineSheet = workbook.createSheet("HIDDEN_MACHINE");
            int rowIndexSheet = 0;
            for (String curingWCT : machineCuringWCT) {
                Row row = hiddenMachineSheet.createRow(rowIndexSheet++);
                row.createCell(0).setCellValue(curingWCT);
            }
            for (String tassWCT : machineTassWCT) {
                Row row = hiddenMachineSheet.createRow(rowIndexSheet++);
                row.createCell(0).setCellValue(tassWCT);
            }

            // Define named range
            Name namedRangeMachine = workbook.createName();
            namedRangeMachine.setNameName("MachineWCT");
            namedRangeMachine.setRefersToFormula("HIDDEN_MACHINE!$A$1:$A$" + rowIndexSheet);

            // Hide the hidden sheet
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenMachineSheet), true);

            int rowIndex = 1;
            for (MachineAllowence m : machineAllowences) {
                Row dataRow = sheet.createRow(rowIndex++);
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(m.getMACHINE_ALLOW_ID() != null ? m.getMACHINE_ALLOW_ID().doubleValue() : 0.0);
                idCell.setCellStyle(borderStyle);

                Cell idMachineCell = dataRow.createCell(2);
                idMachineCell.setCellValue(m.getID_MACHINE() != null ? m.getID_MACHINE() : "");
                idMachineCell.setCellStyle(borderStyle);

                Cell personResponsibleCell = dataRow.createCell(3);
                personResponsibleCell.setCellValue(m.getPERSON_RESPONSIBLE() != null ? m.getPERSON_RESPONSIBLE() : "");
                personResponsibleCell.setCellStyle(borderStyle);

                Cell shift1Cell = dataRow.createCell(4);
                shift1Cell.setCellValue(m.getSHIFT_1() != null ? m.getSHIFT_1().doubleValue() : 0.0);
                shift1Cell.setCellStyle(borderStyle);

                Cell shift2Cell = dataRow.createCell(5);
                shift2Cell.setCellValue(m.getSHIFT_2() != null ? m.getSHIFT_2().doubleValue() : 0.0);
                shift2Cell.setCellStyle(borderStyle);
                
                Cell shift3Cell = dataRow.createCell(6);
                shift3Cell.setCellValue(m.getSHIFT_3() != null ? m.getSHIFT_3().doubleValue() : 0.0);
                shift3Cell.setCellStyle(borderStyle);

                Cell shift1FridayCell = dataRow.createCell(7);
                shift1FridayCell.setCellValue(m.getSHIFT_1_FRIDAY() != null ? m.getSHIFT_1_FRIDAY().doubleValue() : 0.0);
                shift1FridayCell.setCellStyle(borderStyle);

                Cell totalShiftCell = dataRow.createCell(8);
                totalShiftCell.setCellValue(m.getTOTAL_SHIFT_123() != null ? m.getTOTAL_SHIFT_123().doubleValue() : 0.0);
                totalShiftCell.setCellStyle(borderStyle);
            }

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("MachineWCT");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Machine Allowence");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }


    public ByteArrayInputStream layoutMachineAllowencesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel( );
        return byteArrayInputStream;
    }

    public ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "MACHINE_ALLOW_ID",
            "ID_MACHINE",
            "PERSON_RESPONSIBLE",
            "SHIFT_1",
            "SHIFT_2",
            "SHIFT_3",
            "SHIFT_1_FRIDAY",
            "TOTAL_SHIFT_123"
        };

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            // Fetch the active machine data
            List<MachineCuring> activeMachineCurings = machineCuringRepo.findMachineCuringActive();
            List<MachineTass> activeMachineTasss = machineTassRepo.findMachineTassActive();
            List<String> machineCuringWCT = activeMachineCurings.stream()
                .map(MachineCuring::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());
            List<String> machineTassWCT = activeMachineTasss.stream()
                .map(MachineTass::getWORK_CENTER_TEXT)
                .collect(Collectors.toList());

            // Create the main sheet for MACHINE ALLOWANCE DATA
            Sheet sheet = workbook.createSheet("MACHINE ALLOWANCE DATA");
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            // Cell style with border
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setAlignment(HorizontalAlignment.CENTER);

            // Header style
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Set column widths
            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            // Create empty rows
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }

            // Create hidden sheet
            Sheet hiddenMachineSheet = workbook.createSheet("HIDDEN_MACHINES");
            int rowIndexSheet = 0;
            for (String curingWCT : machineCuringWCT) {
                Row row = hiddenMachineSheet.createRow(rowIndexSheet++);
                row.createCell(0).setCellValue(curingWCT);
            }
            for (String tassWCT : machineTassWCT) {
                Row row = hiddenMachineSheet.createRow(rowIndexSheet++);
                row.createCell(0).setCellValue(tassWCT);
            }

            // Define named range
            Name namedRangeMachine = workbook.createName();
            namedRangeMachine.setNameName("MachineWCT");
            namedRangeMachine.setRefersToFormula("HIDDEN_MACHINES!$A$1:$A$" + rowIndexSheet);

            // Hide the hidden sheet
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenMachineSheet), true);

            // Add data validation
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createFormulaListConstraint("MachineWCT");
            CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, 2, 2); // Column 2 (C) range
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            // Write to output stream
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        }
    }




}
