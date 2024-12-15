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
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Pattern;
import sri.sysint.sri_starter_back.repository.PatternRepo;


@Service
@Transactional
public class PatternServiceImpl {
	@Autowired
	private PatternRepo patternRepo;
	
	public PatternServiceImpl(PatternRepo patternRepo){
        this.patternRepo = patternRepo;
    }
    
    public BigDecimal getNewId() {
        return patternRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<Pattern> getAllPattern() {
        Iterable<Pattern> patterns = patternRepo.getDataOrderId();
        List<Pattern> patternList = new ArrayList<>();
        for (Pattern item : patterns) {
            Pattern patternTemp = new Pattern(item);
            patternList.add(patternTemp);
        }
        
        return patternList;
    }
    
    public Optional<Pattern> getPatternById(BigDecimal id) {
        Optional<Pattern> pattern = patternRepo.findById(id);
        return pattern;
    }
    
    public Pattern savePattern(Pattern pattern) {
        try {
            pattern.setPATTERN_ID(getNewId());
            pattern.setSTATUS(BigDecimal.valueOf(1));
            pattern.setCREATION_DATE(new Date());
            pattern.setLAST_UPDATE_DATE(new Date());
            return patternRepo.save(pattern);
        } catch (Exception e) {
            System.err.println("Error saving pattern: " + e.getMessage());
            throw e;
        }
    }
    
    public Pattern updatePattern(Pattern pattern) {
        try {
            Optional<Pattern> currentPatternOpt = patternRepo.findById(pattern.getPATTERN_ID());
            
            if (currentPatternOpt.isPresent()) {
                Pattern currentPattern = currentPatternOpt.get();
                
                currentPattern.setPATTERN_NAME(pattern.getPATTERN_NAME());
                currentPattern.setLAST_UPDATE_DATE(new Date());
                currentPattern.setLAST_UPDATED_BY(pattern.getLAST_UPDATED_BY());
                
                return patternRepo.save(currentPattern);
            } else {
                throw new RuntimeException("Pattern with ID " + pattern.getPATTERN_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating pattern: " + e.getMessage());
            throw e;
        }
    }
    
    public Pattern deletePattern(Pattern pattern) {
        try {
            Optional<Pattern> currentPatternOpt = patternRepo.findById(pattern.getPATTERN_ID());
            
            if (currentPatternOpt.isPresent()) {
                Pattern currentPattern = currentPatternOpt.get();
                
                currentPattern.setSTATUS(BigDecimal.valueOf(0));
                currentPattern.setLAST_UPDATE_DATE(new Date());
                currentPattern.setLAST_UPDATED_BY(pattern.getLAST_UPDATED_BY());
                
                return patternRepo.save(currentPattern);
            } else {
                throw new RuntimeException("Pattern with ID " + pattern.getPATTERN_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating pattern: " + e.getMessage());
            throw e;
        }
    }
    
    public Pattern restorePattern(Pattern pattern) {
        try {
            Optional<Pattern> currentPatternOpt = patternRepo.findById(pattern.getPATTERN_ID());
            
            if (currentPatternOpt.isPresent()) {
                Pattern currentPattern = currentPatternOpt.get();
                
                currentPattern.setSTATUS(BigDecimal.valueOf(1)); 
                currentPattern.setLAST_UPDATE_DATE(new Date());
                currentPattern.setLAST_UPDATED_BY(pattern.getLAST_UPDATED_BY());
                
                return patternRepo.save(currentPattern);
            } else {
                throw new RuntimeException("Pattern with ID " + pattern.getPATTERN_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring pattern: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllPattern() {
        patternRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportPatternsExcel() throws IOException {
        List<Pattern> patterns = patternRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(patterns);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<Pattern> patterns) throws IOException {
        String[] header = {
            "NOMOR",
            "PATTERN_ID",
            "PATTERN_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PATTERN DATA");
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
            for (Pattern p : patterns) {
                Row dataRow = sheet.createRow(rowIndex++);
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell patternIdCell = dataRow.createCell(1);
                patternIdCell.setCellValue(p.getPATTERN_ID() != null ? p.getPATTERN_ID().doubleValue() : null);
                patternIdCell.setCellStyle(borderStyle);

                Cell patternNameCell = dataRow.createCell(2);
                patternNameCell.setCellValue(p.getPATTERN_NAME() != null ? p.getPATTERN_NAME() : "");
                patternNameCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Pattern");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutPatternsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "PATTERN_ID",
            "PATTERN_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PATTERN DATA");
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

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }


            int rowIndex = 1;
            
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Pattern");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
}