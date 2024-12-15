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
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;


@Service
@Transactional
public class ItemCuringServiceImpl {
	@Autowired
    private ItemCuringRepo itemCuringRepo;

    public ItemCuringServiceImpl(ItemCuringRepo itemCuringRepo) {
        this.itemCuringRepo = itemCuringRepo;
    }

    public List<ItemCuring> getAllItemCuring() {
        Iterable<ItemCuring> itemCurings = itemCuringRepo.getDataOrderId();
        List<ItemCuring> itemCuringList = new ArrayList<>();
        for (ItemCuring item : itemCurings) {
            ItemCuring itemCuringTemp = new ItemCuring(item);
            itemCuringList.add(itemCuringTemp);
        }
        return itemCuringList;
    }

    public Optional<ItemCuring> getItemCuringById(String id) {
        Optional<ItemCuring> itemCuring = itemCuringRepo.findById(id);
        return itemCuring;
    }

    public ItemCuring saveItemCuring(ItemCuring itemCuring) {
        try {
            itemCuring.setSTATUS(BigDecimal.valueOf(1));
            itemCuring.setCREATION_DATE(new Date());
            itemCuring.setLAST_UPDATE_DATE(new Date());
            return itemCuringRepo.save(itemCuring);
        } catch (Exception e) {
            System.err.println("Error saving itemCuring: " + e.getMessage());
            throw e;
        }
    }

    public ItemCuring updateItemCuring(ItemCuring itemCuring) {
        try {
            Optional<ItemCuring> currentItemCuringOpt = itemCuringRepo.findById(itemCuring.getITEM_CURING());
            if (currentItemCuringOpt.isPresent()) {
                ItemCuring currentItemCuring = currentItemCuringOpt.get();
                currentItemCuring.setKAPA_PER_MOULD(itemCuring.getKAPA_PER_MOULD());
                currentItemCuring.setNUMBER_OF_MOULD(itemCuring.getNUMBER_OF_MOULD());
                currentItemCuring.setMACHINE_TYPE(itemCuring.getMACHINE_TYPE());
                currentItemCuring.setSPARE_MOULD(itemCuring.getSPARE_MOULD());
                currentItemCuring.setMOULD_MONTHLY_PLAN(itemCuring.getMOULD_MONTHLY_PLAN());

                currentItemCuring.setLAST_UPDATE_DATE(new Date());
                currentItemCuring.setLAST_UPDATED_BY(itemCuring.getLAST_UPDATED_BY());
                return itemCuringRepo.save(currentItemCuring);
            } else {
                throw new RuntimeException("ItemCuring with ID " + itemCuring.getITEM_CURING() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating itemCuring: " + e.getMessage());
            throw e;
        }
    }

    public ItemCuring deleteItemCuring(ItemCuring itemCuring) {
        try {
            Optional<ItemCuring> currentItemCuringOpt = itemCuringRepo.findById(itemCuring.getITEM_CURING());
            if (currentItemCuringOpt.isPresent()) {
                ItemCuring currentItemCuring = currentItemCuringOpt.get();
                currentItemCuring.setSTATUS(BigDecimal.valueOf(0));
                currentItemCuring.setLAST_UPDATE_DATE(new Date());
                currentItemCuring.setLAST_UPDATED_BY(itemCuring.getLAST_UPDATED_BY());
                return itemCuringRepo.save(currentItemCuring);
            } else {
                throw new RuntimeException("ItemCuring with ID " + itemCuring.getITEM_CURING() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating itemCuring: " + e.getMessage());
            throw e;
        }
    }
    
    public ItemCuring restoreItemCuring(ItemCuring itemCuring) {
        try {
            Optional<ItemCuring> currentItemCuringOpt = itemCuringRepo.findById(itemCuring.getITEM_CURING());
            
            if (currentItemCuringOpt.isPresent()) {
                ItemCuring currentItemCuring = currentItemCuringOpt.get();
                
                currentItemCuring.setSTATUS(BigDecimal.valueOf(1)); // Mengubah status menjadi 1 untuk restore
                currentItemCuring.setLAST_UPDATE_DATE(new Date());
                currentItemCuring.setLAST_UPDATED_BY(itemCuring.getLAST_UPDATED_BY());
                
                return itemCuringRepo.save(currentItemCuring);
            } else {
                throw new RuntimeException("ItemCuring with ID " + itemCuring.getITEM_CURING() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring itemCuring: " + e.getMessage());
            throw e;
        }
    }

    public void deleteAllItemCuring() {
        itemCuringRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportItemCuringsExcel() throws IOException {
        List<ItemCuring> itemCurings = itemCuringRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(itemCurings);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<ItemCuring> itemCurings) throws IOException {
        String[] header = {"NOMOR", "ITEM_CURING", "KAPA_PER_MOULD", "NUMBER_OF_MOULD", "MACHINE_TYPE", "SPARE_MOULD", "MOULD_MONTHLY_PLAN"};

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("ITEM CURING DATA");

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
            for (ItemCuring item : itemCurings) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(item.getITEM_CURING());
                idCell.setCellStyle(borderStyle);

                Cell kapaCell = dataRow.createCell(2);
                kapaCell.setCellValue(item.getKAPA_PER_MOULD() != null ? item.getKAPA_PER_MOULD().doubleValue() : null);
                kapaCell.setCellStyle(borderStyle);

                Cell numberOfMouldCell = dataRow.createCell(3);
                numberOfMouldCell.setCellValue(item.getNUMBER_OF_MOULD() != null ? item.getNUMBER_OF_MOULD().doubleValue() : null);
                numberOfMouldCell.setCellStyle(borderStyle);

                Cell machineTypeCell = dataRow.createCell(4);
                machineTypeCell.setCellValue(item.getMACHINE_TYPE() != null ? item.getMACHINE_TYPE() : "");
                machineTypeCell.setCellStyle(borderStyle);

                Cell spareMouldCell = dataRow.createCell(5);
                spareMouldCell.setCellValue(item.getSPARE_MOULD() != null ? item.getSPARE_MOULD().doubleValue() : null);
                spareMouldCell.setCellStyle(borderStyle);

                Cell mouldPlanCell = dataRow.createCell(6);
                mouldPlanCell.setCellValue(item.getMOULD_MONTHLY_PLAN() != null ? item.getMOULD_MONTHLY_PLAN().doubleValue() : 0);
                mouldPlanCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            System.err.println("Failed to export ItemCuring data: " + e.getMessage());
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
    
    public ByteArrayInputStream layoutItemCuringsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {"NOMOR", "ITEM_CURING", "KAPA_PER_MOULD", "NUMBER_OF_MOULD", "MACHINE_TYPE", "SPARE_MOULD", "MOULD_MONTHLY_PLAN"};

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("ITEM CURING DATA");

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
            System.err.println("Failed to export ItemCuring data: " + e.getMessage());
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}