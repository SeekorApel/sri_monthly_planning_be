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
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.ProductType;
import sri.sysint.sri_starter_back.repository.ProductTypeRepo;

@Service
@Transactional
public class ProducTypeServiceImpl {
	@Autowired
    private ProductTypeRepo productTypeRepo;
	
	public ProducTypeServiceImpl(ProductTypeRepo productTypeRepo){
        this.productTypeRepo = productTypeRepo;
    }
    
    public BigDecimal getNewId() {
    	return productTypeRepo.getNewId().add(BigDecimal.valueOf(1));
    }
    
    public List<ProductType> getAllProductType() {
    	Iterable<ProductType> productTypes = productTypeRepo.getDataOrderId();
        List<ProductType> productTypeList = new ArrayList<>();
        for (ProductType item : productTypes) {
            ProductType productTypeTemp = new ProductType(item);
            productTypeList.add(productTypeTemp);
        }
        
        return productTypeList;
    }
    
    public Optional<ProductType> getProductTypeById(BigDecimal id) {
    	Optional<ProductType> productType = productTypeRepo.findById(id);
    	return productType;
    }
    
    public ProductType saveProductType(ProductType productType) {
        try {
        	productType.setPRODUCT_TYPE_ID(getNewId());
        	productType.setSTATUS(BigDecimal.valueOf(1));
        	productType.setCREATION_DATE(new Date());
        	productType.setLAST_UPDATE_DATE(new Date());
            return productTypeRepo.save(productType);
        } catch (Exception e) {
            System.err.println("Error saving plant: " + e.getMessage());
            throw e;
        }
    }
    

    public ProductType updateProductType(ProductType productType) {
        try {
            Optional<ProductType> currentProductTypeOpt = productTypeRepo.findById(productType.getPRODUCT_TYPE_ID());
            
            if (currentProductTypeOpt.isPresent()) {
            	ProductType currentProductType = currentProductTypeOpt.get();
                
            	currentProductType.setPRODUCT_MERK(productType.getPRODUCT_MERK());
            	currentProductType.setPRODUCT_TYPE(productType.getPRODUCT_TYPE());
            	currentProductType.setCATEGORY(productType.getCATEGORY());
            	currentProductType.setLAST_UPDATE_DATE(new Date());
            	currentProductType.setLAST_UPDATED_BY(productType.getLAST_UPDATED_BY());
                
                return productTypeRepo.save(currentProductType);
            } else {
                throw new RuntimeException("Plant with ID " + productType.getPRODUCT_TYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating plant: " + e.getMessage());
            throw e;
        }
    }
    
    public ProductType deleteProductType(ProductType productType) {
        try {
            Optional<ProductType> currentProductTypeOpt = productTypeRepo.findById(productType.getPRODUCT_TYPE_ID());
            
            if (currentProductTypeOpt.isPresent()) {
            	ProductType currentProductType = currentProductTypeOpt.get();
                
            	currentProductType.setSTATUS(BigDecimal.valueOf(0));
            	currentProductType.setLAST_UPDATE_DATE(new Date());
            	currentProductType.setLAST_UPDATED_BY(productType.getLAST_UPDATED_BY());
                
                return productTypeRepo.save(currentProductType);
            } else {
                throw new RuntimeException("Plant with ID " + productType.getPRODUCT_TYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating plant: " + e.getMessage());
            throw e;
        }
    }
    public ProductType activateProductType(ProductType productType) {
        try {
            Optional<ProductType> currentProductTypeOpt = productTypeRepo.findById(productType.getPRODUCT_TYPE_ID());
            
            if (currentProductTypeOpt.isPresent()) {
            	ProductType currentProductType = currentProductTypeOpt.get();
                
            	currentProductType.setSTATUS(BigDecimal.valueOf(1));
            	currentProductType.setLAST_UPDATE_DATE(new Date());
            	currentProductType.setLAST_UPDATED_BY(productType.getLAST_UPDATED_BY());
                
                return productTypeRepo.save(currentProductType);
            } else {
                throw new RuntimeException("Plant with ID " + productType.getPRODUCT_TYPE_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating plant: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllProductType() {
    	productTypeRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportProductTypesExcel() throws IOException {
        List<ProductType> productTypes = productTypeRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(productTypes);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<ProductType> productTypes) throws IOException {
        String[] header = {
            "NOMOR",
            "PRODUCT_TYPE_ID",
            "PRODUCT_MERK",
            "PRODUCT_TYPE",
            "CATEGORY"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PRODUCT TYPE DATA");
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
            for (ProductType p : productTypes) {
                Row dataRow = sheet.createRow(rowIndex++);
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(p.getPRODUCT_TYPE_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell productMerkCell = dataRow.createCell(2);
                productMerkCell.setCellValue(p.getPRODUCT_MERK() != null ? p.getPRODUCT_MERK() : null);
                productMerkCell.setCellStyle(borderStyle);
                
                Cell productTypeCell = dataRow.createCell(3);
                productTypeCell.setCellValue(p.getPRODUCT_TYPE() != null ? p.getPRODUCT_TYPE() : null);
                productTypeCell.setCellStyle(borderStyle);
                
                Cell categoryCell = dataRow.createCell(4);
                categoryCell.setCellValue(p.getCATEGORY() != null ? p.getCATEGORY().toString() : "");
                categoryCell.setCellStyle(borderStyle);
            }
            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Product type");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
    
    public ByteArrayInputStream layoutProductTypesExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "PRODUCT_TYPE_ID",
            "PRODUCT_MERK",
            "PRODUCT_TYPE",
            "CATEGORY"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PRODUCT TYPE DATA");
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
            System.out.println("Gagal mengekspor data Product type");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
