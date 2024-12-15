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
import sri.sysint.sri_starter_back.model.ItemAssy;
import sri.sysint.sri_starter_back.repository.ItemAssyRepo;


@Service
@Transactional
public class ItemAssyServiceImpl {
    @Autowired
    private ItemAssyRepo itemAssyRepo;

    public ItemAssyServiceImpl(ItemAssyRepo itemAssyRepo) {
        this.itemAssyRepo = itemAssyRepo;
    }

    public List<ItemAssy> getAllItemAssy() {
        Iterable<ItemAssy> itemAssys = itemAssyRepo.getDataOrderId();
        List<ItemAssy> itemAssyList = new ArrayList<>();
        for (ItemAssy item : itemAssys) {
            ItemAssy itemAssyTemp = new ItemAssy(item);
            itemAssyList.add(itemAssyTemp);
        }
        return itemAssyList;
    }

    public Optional<ItemAssy> getItemAssyById(String id) {
        Optional<ItemAssy> itemAssy = itemAssyRepo.findById(id);
        return itemAssy;
    }

    public ItemAssy saveItemAssy(ItemAssy itemAssy) {
        try {
        	itemAssy.setITEM_ASSY(itemAssy.getITEM_ASSY());
            itemAssy.setSTATUS(BigDecimal.valueOf(1));
            itemAssy.setCREATION_DATE(new Date());
            itemAssy.setLAST_UPDATE_DATE(new Date());
            return itemAssyRepo.save(itemAssy);
        } catch (Exception e) {
            System.err.println("Error saving itemAssy: " + e.getMessage());
            throw e;
        }
    }

    public ItemAssy updateItemAssy(ItemAssy itemAssy) {
        try {
            Optional<ItemAssy> currentItemAssyOpt = itemAssyRepo.findById(itemAssy.getITEM_ASSY());
            if (currentItemAssyOpt.isPresent()) {
                ItemAssy currentItemAssy = currentItemAssyOpt.get();
                currentItemAssy.setLAST_UPDATE_DATE(new Date());
                currentItemAssy.setLAST_UPDATED_BY(itemAssy.getLAST_UPDATED_BY());
                return itemAssyRepo.save(currentItemAssy);
            } else {
                throw new RuntimeException("ItemAssy with ID " + itemAssy.getITEM_ASSY() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating itemAssy: " + e.getMessage());
            throw e;
        }
    }

    public ItemAssy deleteItemAssy(ItemAssy itemAssy) {
        try {
            Optional<ItemAssy> currentItemAssyOpt = itemAssyRepo.findById(itemAssy.getITEM_ASSY());
            if (currentItemAssyOpt.isPresent()) {
                ItemAssy currentItemAssy = currentItemAssyOpt.get();
                currentItemAssy.setSTATUS(BigDecimal.valueOf(0));
                currentItemAssy.setLAST_UPDATE_DATE(new Date());
                currentItemAssy.setLAST_UPDATED_BY(itemAssy.getLAST_UPDATED_BY());
                return itemAssyRepo.save(currentItemAssy);
            } else {
                throw new RuntimeException("ItemAssy with ID " + itemAssy.getITEM_ASSY() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating itemAssy: " + e.getMessage());
            throw e;
        }
    }
    
    public ItemAssy restoreItemAssy(ItemAssy itemAssy) {
        try {
            Optional<ItemAssy> currentItemAssyOpt = itemAssyRepo.findById(itemAssy.getITEM_ASSY());
            if (currentItemAssyOpt.isPresent()) {
                ItemAssy currentItemAssy = currentItemAssyOpt.get();
                currentItemAssy.setSTATUS(BigDecimal.valueOf(1)); 
                currentItemAssy.setLAST_UPDATE_DATE(new Date());
                currentItemAssy.setLAST_UPDATED_BY(itemAssy.getLAST_UPDATED_BY());
                return itemAssyRepo.save(currentItemAssy);
            } else {
                throw new RuntimeException("ItemAssy with ID " + itemAssy.getITEM_ASSY() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring itemAssy: " + e.getMessage());
            throw e;
        }
    }


    public void deleteAllItemAssy() {
        itemAssyRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportItemAssysExcel() throws IOException {
        List<ItemAssy> itemAssys = itemAssyRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(itemAssys);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<ItemAssy> itemAssys) throws IOException {
        String[] header = {"NOMOR", "ITEM_ASSY"};

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("ITEM ASSY DATA");

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

            sheet.setColumnWidth(0, 10 * 256);
            sheet.setColumnWidth(1, 20 * 256);

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 1;
            int nomor = 1;
            for (ItemAssy item : itemAssys) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell itemAssyCell = dataRow.createCell(1);
                itemAssyCell.setCellValue(item.getITEM_ASSY());
                itemAssyCell.setCellStyle(borderStyle);
            }
            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export ItemAssy data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutItemAssysExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {"NOMOR", "ITEM_ASSY"};

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("ITEM ASSY DATA");

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

            sheet.setColumnWidth(0, 10 * 256);
            sheet.setColumnWidth(1, 20 * 256);

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
            int nomor = 1;
            
            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export ItemAssy data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
}
