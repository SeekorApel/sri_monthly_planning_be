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
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.repository.SizeRepo;


@Service
@Transactional
public class SizeServiceImpl {
	@Autowired
	private SizeRepo sizeRepo;

	public SizeServiceImpl(SizeRepo sizeRepo) {
        this.sizeRepo = sizeRepo;
    }
	
    public List<Size> getAllSize() {
        Iterable<Size> sizes = sizeRepo.findAll();
        List<Size> sizeList = new ArrayList<>();
        for (Size item : sizes) {
            Size sizeTemp = new Size(item);
            sizeList.add(sizeTemp);
        }
        return sizeList;
    }

    public Optional<Size> getSizeById(String id) {
        Optional<Size> size = sizeRepo.findById(id);
        return size;
    }

    public Size saveSize(Size size) {
        try {
            size.setSIZE_ID(size.getSIZE_ID());
            size.setSTATUS(BigDecimal.valueOf(1));
            size.setCREATION_DATE(new Date());
            size.setLAST_UPDATE_DATE(new Date());
            return sizeRepo.save(size);
        } catch (Exception e) {
            System.err.println("Error saving size: " + e.getMessage());
            throw e;
        }
    }

    public Size updateSize(Size size) {
        try {
            Optional<Size> currentSizeOpt = sizeRepo.findById(size.getSIZE_ID());
            if (currentSizeOpt.isPresent()) {
                Size currentSize = currentSizeOpt.get();
                currentSize.setDESCRIPTION(size.getDESCRIPTION());
                currentSize.setLAST_UPDATE_DATE(new Date());
                currentSize.setLAST_UPDATED_BY(size.getLAST_UPDATED_BY());
                return sizeRepo.save(currentSize);
            } else {
                throw new RuntimeException("Size with ID " + size.getSIZE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating size: " + e.getMessage());
            throw e;
        }
    }

    public Size deleteSize(Size size) {
        try {
            Optional<Size> currentSizeOpt = sizeRepo.findById(size.getSIZE_ID());
            if (currentSizeOpt.isPresent()) {
                Size currentSize = currentSizeOpt.get();
                currentSize.setSTATUS(BigDecimal.valueOf(0));
                currentSize.setLAST_UPDATE_DATE(new Date());
                currentSize.setLAST_UPDATED_BY(size.getLAST_UPDATED_BY());
                return sizeRepo.save(currentSize);
            } else {
                throw new RuntimeException("Size with ID " + size.getSIZE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating size: " + e.getMessage());
            throw e;
        }
    }
    
    public Size restoreSize(Size size) {
        try {
            Optional<Size> currentSizeOpt = sizeRepo.findById(size.getSIZE_ID());
            if (currentSizeOpt.isPresent()) {
                Size currentSize = currentSizeOpt.get();
                currentSize.setSTATUS(BigDecimal.valueOf(1)); 
                currentSize.setLAST_UPDATE_DATE(new Date());
                currentSize.setLAST_UPDATED_BY(size.getLAST_UPDATED_BY());
                return sizeRepo.save(currentSize);
            } else {
                throw new RuntimeException("Size with ID " + size.getSIZE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring size: " + e.getMessage());
            throw e;
        }
    }


    public void deleteAllSize() {
        sizeRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportSizesExcel() throws IOException {
        List<Size> sizes = sizeRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(sizes);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<Size> sizes) throws IOException {
        String[] header = {
            "NOMOR",
            "SIZE_ID",
            "DESCRIPTION"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("SIZE DATA");
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
            for (Size s : sizes) {
                Row dataRow = sheet.createRow(rowIndex++);
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(s.getSIZE_ID());
                idCell.setCellStyle(borderStyle);

                Cell descriptionCell = dataRow.createCell(2);
                descriptionCell.setCellValue(s.getDESCRIPTION() != null ? s.getDESCRIPTION() : "");
                descriptionCell.setCellStyle(borderStyle);
            }

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export Size data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutSizesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "SIZE_ID",
            "DESCRIPTION"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("SIZE DATA");
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
            System.out.println("Failed to export Size data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
    
}