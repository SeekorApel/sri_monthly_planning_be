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
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Setting;
import sri.sysint.sri_starter_back.repository.SettingRepo;


@Service
@Transactional
public class SettingServiceImpl {
	@Autowired
	private SettingRepo settingRepo;
	
	public SettingServiceImpl(SettingRepo settingRepo) {
        this.settingRepo = settingRepo;
    }
    
	public BigDecimal getNewId() {
        return settingRepo.getNewId().add(BigDecimal.valueOf(1));
    }

	public List<Setting> getAllSettings() {
        Iterable<Setting> settings = settingRepo.getDataOrderId();
        List<Setting> settingList = new ArrayList<>();
        for (Setting item : settings) {
            Setting settingTemp = new Setting(item);
            settingList.add(settingTemp);
        }
        return settingList;
    }
    
	public Optional<Setting> getSettingById(BigDecimal id) {
        Optional<Setting> setting = settingRepo.findById(id);
        return setting;
    }
    
    public Setting saveSetting(Setting setting) {
        try {
        	String settingValue = setting.getSETTING_VALUE();
            if (settingValue != null && settingValue.contains(".")) {
                if (!settingValue.contains(",")) {
                    settingValue = settingValue.replace('.', ',');
                }
            }
            setting.setSETTING_ID(getNewId());
            setting.setSTATUS(BigDecimal.valueOf(1));
            setting.setSETTING_VALUE(settingValue);
            setting.setCREATION_DATE(new Date());
            setting.setLAST_UPDATE_DATE(new Date());
            return settingRepo.save(setting);
        } catch (Exception e) {
            System.err.println("Error saving setting: " + e.getMessage());
            throw e;
        }
    }
    
    public Setting updateSetting(Setting setting) {
        try {
            Optional<Setting> currentSettingOpt = settingRepo.findById(setting.getSETTING_ID());
            if (currentSettingOpt.isPresent()) {
                Setting currentSetting = currentSettingOpt.get();
                
                String settingValue = setting.getSETTING_VALUE();
                if (settingValue != null && settingValue.contains(".")) {
                    if (!settingValue.contains(",")) {
                        settingValue = settingValue.replace('.', ',');
                    }
                }
                currentSetting.setSETTING_KEY(setting.getSETTING_KEY());
                currentSetting.setSETTING_VALUE(settingValue);
                currentSetting.setDESCRIPTION(setting.getDESCRIPTION());
                currentSetting.setLAST_UPDATE_DATE(new Date());
                currentSetting.setLAST_UPDATED_BY(setting.getLAST_UPDATED_BY());
                return settingRepo.save(currentSetting);
            } else {
                throw new RuntimeException("Setting with ID " + setting.getSETTING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating setting: " + e.getMessage());
            throw e;
        }
    }
    
    public Setting deleteSetting(Setting setting) {
        try {
            Optional<Setting> currentSettingOpt = settingRepo.findById(setting.getSETTING_ID());
            if (currentSettingOpt.isPresent()) {
                Setting currentSetting = currentSettingOpt.get();
                currentSetting.setSTATUS(BigDecimal.valueOf(0));
                currentSetting.setLAST_UPDATE_DATE(new Date());
                currentSetting.setLAST_UPDATED_BY(setting.getLAST_UPDATED_BY());
                return settingRepo.save(currentSetting);
            } else {
                throw new RuntimeException("Setting with ID " + setting.getSETTING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating setting: " + e.getMessage());
            throw e;
        }
    }
    
    public Setting restoreSetting(Setting setting) {
        try {
            Optional<Setting> currentSettingOpt = settingRepo.findById(setting.getSETTING_ID());
            
            if (currentSettingOpt.isPresent()) {
                Setting currentSetting = currentSettingOpt.get();
                
                currentSetting.setSTATUS(BigDecimal.valueOf(1)); 
                currentSetting.setLAST_UPDATE_DATE(new Date());
                currentSetting.setLAST_UPDATED_BY(setting.getLAST_UPDATED_BY());
                
                return settingRepo.save(currentSetting);
            } else {
                throw new RuntimeException("Setting with ID " + setting.getSETTING_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring setting: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllSettings() {
        settingRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportSettingsExcel() throws IOException {
        List<Setting> settings = settingRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(settings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<Setting> settings) throws IOException {
        String[] header = {
            "NOMOR",
            "SETTING_ID",
            "SETTING_KEY",
            "SETTING_VALUE",
            "DESCRIPTION"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("SETTING DATA");
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
            borderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

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
            for (Setting s : settings) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(s.getSETTING_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell settingKeyCell = dataRow.createCell(2);
                settingKeyCell.setCellValue(s.getSETTING_KEY() != null ? s.getSETTING_KEY() : null);
                settingKeyCell.setCellStyle(borderStyle);

                Cell settingValueCell = dataRow.createCell(3);
                settingValueCell.setCellValue(s.getSETTING_VALUE() != null ? s.getSETTING_VALUE() : null);
                settingValueCell.setCellStyle(borderStyle);

                Cell descriptionCell = dataRow.createCell(4);
                descriptionCell.setCellValue(s.getDESCRIPTION() != null ? s.getDESCRIPTION() : null);
                descriptionCell.setCellStyle(borderStyle);
            }

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Setting");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
    
    public ByteArrayInputStream layoutSettingsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "SETTING_ID",
            "SETTING_KEY",
            "SETTING_VALUE",
            "DESCRIPTION"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("SETTING DATA");
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
            borderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

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
            
            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Setting");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }


}