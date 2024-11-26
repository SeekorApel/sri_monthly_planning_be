
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

import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;


@Service
@Transactional
public class MachineCuringServiceImpl {
    @Autowired
    private MachineCuringRepo machineCuringRepo;

    public MachineCuringServiceImpl(MachineCuringRepo machineCuringRepo) {
        this.machineCuringRepo = machineCuringRepo;
    }

    public List<MachineCuring> getAllMachineCuring() {
        Iterable<MachineCuring> machineCurings = machineCuringRepo.getDataOrderId();
        List<MachineCuring> machineCuringList = new ArrayList<>();
        for (MachineCuring machine : machineCurings) {
            MachineCuring machineCuringTemp = new MachineCuring(machine);
            machineCuringList.add(machineCuringTemp);
        }
        return machineCuringList;
    }

    public Optional<MachineCuring> getMachineCuringById(String id) {
        Optional<MachineCuring> machineCuring = machineCuringRepo.findById(id);
        return machineCuring;
    }

    public MachineCuring saveMachineCuring(MachineCuring machineCuring) {
        try {
            machineCuring.setWORK_CENTER_TEXT(machineCuring.getWORK_CENTER_TEXT());
            machineCuring.setSTATUS(BigDecimal.valueOf(1));
            machineCuring.setCREATION_DATE(new Date());
            machineCuring.setLAST_UPDATE_DATE(new Date());
            return machineCuringRepo.save(machineCuring);
        } catch (Exception e) {
            System.err.println("Error saving machineCuring: " + e.getMessage());
            throw e;
        }
    }

    public MachineCuring updateMachineCuring(MachineCuring machineCuring) {
        try {
            Optional<MachineCuring> currentMachineCuringOpt = machineCuringRepo.findById(machineCuring.getWORK_CENTER_TEXT());
            if (currentMachineCuringOpt.isPresent()) {
                MachineCuring currentMachineCuring = currentMachineCuringOpt.get();
                currentMachineCuring.setLAST_UPDATE_DATE(new Date());
                currentMachineCuring.setBUILDING_ID(machineCuring.getBUILDING_ID());
                currentMachineCuring.setCAVITY(machineCuring.getCAVITY());
                currentMachineCuring.setMACHINE_TYPE(machineCuring.getMACHINE_TYPE());
                currentMachineCuring.setSTATUS_USAGE(machineCuring.getSTATUS_USAGE());


                currentMachineCuring.setLAST_UPDATED_BY(machineCuring.getLAST_UPDATED_BY());
                return machineCuringRepo.save(currentMachineCuring);
            } else {
                throw new RuntimeException("MachineCuring with ID " + machineCuring.getWORK_CENTER_TEXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineCuring: " + e.getMessage());
            throw e;
        }
    }

    public MachineCuring deleteMachineCuring(MachineCuring machineCuring) {
        try {
            Optional<MachineCuring> currentMachineCuringOpt = machineCuringRepo.findById(machineCuring.getWORK_CENTER_TEXT());
            if (currentMachineCuringOpt.isPresent()) {
                MachineCuring currentMachineCuring = currentMachineCuringOpt.get();
                currentMachineCuring.setSTATUS(BigDecimal.valueOf(0));
                currentMachineCuring.setLAST_UPDATE_DATE(new Date());
                currentMachineCuring.setLAST_UPDATED_BY(machineCuring.getLAST_UPDATED_BY());
                return machineCuringRepo.save(currentMachineCuring);
            } else {
                throw new RuntimeException("MachineCuring with ID " + machineCuring.getWORK_CENTER_TEXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating machineCuring: " + e.getMessage());
            throw e;
        }
    }
    
    public MachineCuring restoreMachineCuring(MachineCuring machineCuring) {
        try {
            Optional<MachineCuring> currentMachineCuringOpt = machineCuringRepo.findById(machineCuring.getWORK_CENTER_TEXT());
            
            if (currentMachineCuringOpt.isPresent()) {
                MachineCuring currentMachineCuring = currentMachineCuringOpt.get();
                
                currentMachineCuring.setSTATUS(BigDecimal.valueOf(1)); 
                currentMachineCuring.setLAST_UPDATE_DATE(new Date());
                currentMachineCuring.setLAST_UPDATED_BY(machineCuring.getLAST_UPDATED_BY());
                
                return machineCuringRepo.save(currentMachineCuring);
            } else {
                throw new RuntimeException("MachineCuring with ID " + machineCuring.getWORK_CENTER_TEXT() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring machineCuring: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllMachineCuring() {
        machineCuringRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportMachineCuringsExcel() throws IOException {
        List<MachineCuring> machineCurings = machineCuringRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(machineCurings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<MachineCuring> machineCurings) throws IOException {
        String[] header = {
            "NOMOR", 
            "WORK_CENTER_TEXT",
            "BUILDING_ID",
            "CAVITY",
            "MACHINE_TYPE",
            "STATUS_USAGE"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("MACHINE CURING DATA");

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
            int nomor = 1;
            for (MachineCuring mc : machineCurings) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell workcentertextcell = dataRow.createCell(1);
                workcentertextcell.setCellValue(mc.getWORK_CENTER_TEXT());
                workcentertextcell.setCellStyle(borderStyle);

                Cell buildingIdCell = dataRow.createCell(2);
                buildingIdCell.setCellValue(mc.getBUILDING_ID().doubleValue());
                buildingIdCell.setCellStyle(borderStyle);

                Cell cavityCell = dataRow.createCell(3);
                cavityCell.setCellValue(mc.getCAVITY().doubleValue());
                cavityCell.setCellStyle(borderStyle);

                Cell machinetypecell = dataRow.createCell(4);
                machinetypecell.setCellValue(mc.getMACHINE_TYPE());
                machinetypecell.setCellStyle(borderStyle);

                Cell statusUsageCell = dataRow.createCell(5);
                statusUsageCell.setCellValue(mc.getSTATUS_USAGE().doubleValue());
                statusUsageCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            throw new IOException("Failed to export MachineCuring data", e);
        } finally {
            workbook.close();
            out.close();
        }
    }

    
}
