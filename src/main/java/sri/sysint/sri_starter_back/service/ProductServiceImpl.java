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

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.repository.ProductRepo;

@Service
@Transactional
public class ProductServiceImpl {
	
	@Autowired
    private ProductRepo productRepo;
	
    public ProductServiceImpl(ProductRepo productRepo){
        this.productRepo = productRepo;
    }
    
    public BigDecimal getNewId() {
    	return productRepo.getNewId().add(BigDecimal.valueOf(1));
    }
    
    public List<Product> getAllProduct() {
    	Iterable<Product> products = productRepo.getDataOrderId();
        List<Product> productList = new ArrayList<>();
        for (Product item : products) {
        	Product productTemp = new Product(item);
        	productList.add(productTemp);
        }
        
        return productList;
    }
    
    public Optional<Product> getProductById(BigDecimal id) {
    	Optional<Product> product = productRepo.findById(id);
    	return product;
    }
    
    public Product saveProduct(Product product) {
        try {
        	product.setPART_NUMBER(product.getPART_NUMBER());
        	product.setSTATUS(BigDecimal.valueOf(1));
        	product.setCREATION_DATE(new Date());
        	product.setLAST_UPDATE_DATE(new Date());
            return productRepo.save(product);
        } catch (Exception e) {
            System.err.println("Error saving Product: " + e.getMessage());
            throw e;
        }
    }
    
    public Product updateProduct(Product product) {
        try {
            Optional<Product> currentProductOpt = productRepo.findById(product.getPART_NUMBER());
            
            if (currentProductOpt.isPresent()) {
            	Product currentProduct = currentProductOpt.get();
                
            	currentProduct.setITEM_CURING(product.getITEM_CURING());
            	currentProduct.setPATTERN_ID(product.getPATTERN_ID());
            	currentProduct.setSIZE_ID(product.getSIZE_ID());
            	currentProduct.setPRODUCT_TYPE_ID(product.getPRODUCT_TYPE_ID());
            	currentProduct.setDESCRIPTION(product.getDESCRIPTION());
            	currentProduct.setRIM(product.getRIM());
            	currentProduct.setWIB_TUBE(product.getWIB_TUBE());
            	currentProduct.setITEM_ASSY(product.getITEM_ASSY());
            	currentProduct.setITEM_EXT(product.getITEM_EXT());
            	currentProduct.setEXT_DESCRIPTION(product.getEXT_DESCRIPTION());
            	currentProduct.setQTY_PER_RAK(product.getQTY_PER_RAK());
            	currentProduct.setUPPER_CONSTANT(product.getUPPER_CONSTANT());
            	currentProduct.setLOWER_CONSTANT(product.getLOWER_CONSTANT());
            	currentProduct.setLAST_UPDATE_DATE(new Date());
            	currentProduct.setLAST_UPDATED_BY(product.getLAST_UPDATED_BY());
                
                return productRepo.save(currentProduct);
            } else {
                throw new RuntimeException("Product with ID " + product.getPART_NUMBER() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Product: " + e.getMessage());
            throw e;
        }
    }
    
    public Product deleteProduct(Product product) {
        try {
            Optional<Product> currentProductOpt = productRepo.findById(product.getPART_NUMBER());
            
            if (currentProductOpt.isPresent()) {
            	Product currentProduct = currentProductOpt.get();
                
            	currentProduct.setSTATUS(BigDecimal.valueOf(0));
            	currentProduct.setLAST_UPDATE_DATE(new Date());
            	currentProduct.setLAST_UPDATED_BY(product.getLAST_UPDATED_BY());
                
                return productRepo.save(currentProduct);
            } else {
                throw new RuntimeException("Product with ID " + product.getPART_NUMBER() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Product: " + e.getMessage());
            throw e;
        }
    }
    public Product activateProduct(Product product) {
        try {
            Optional<Product> currentProductOpt = productRepo.findById(product.getPART_NUMBER());
            
            if (currentProductOpt.isPresent()) {
            	Product currentProduct = currentProductOpt.get();
                
            	currentProduct.setSTATUS(BigDecimal.valueOf(1));
            	currentProduct.setLAST_UPDATE_DATE(new Date());
            	currentProduct.setLAST_UPDATED_BY(product.getLAST_UPDATED_BY());
                
                return productRepo.save(currentProduct);
            } else {
                throw new RuntimeException("Product with ID " + product.getPART_NUMBER() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating Product: " + e.getMessage());
            throw e;
        }
    }
    
    public void deleteAllProduct() {
    	productRepo.deleteAll();
    }
    
    public ByteArrayInputStream exportProductsExcel() throws IOException {
        List<Product> products = productRepo.getDataOrderId();
        ByteArrayInputStream byteArrayInputStream = dataToExcel(products);
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream dataToExcel(List<Product> products) throws IOException {
        String[] header = {
            "NOMOR",
            "PART_NUMBER",
            "ITEM_CURING",
            "PATTERN_ID",
            "SIZE_ID",
            "PRODUCT_TYPE_ID",
            "DESCRIPTION",
            "RIM",
            "WIB_TUBE",
            "ITEM_ASSY",
            "ITEM_EXT",
            "EXT_DESCRIPTION",
            "QTY_PER_RAK",
            "UPPER_CONSTANT",
            "LOWER_CONSTANT"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PRODUCT DATA");
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
            for (Product p : products) {
                Row dataRow = sheet.createRow(rowIndex++);
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                Cell partNumberCell = dataRow.createCell(1);
                partNumberCell.setCellValue(p.getPART_NUMBER().doubleValue());
                partNumberCell.setCellStyle(borderStyle);

                Cell itemCuringCell = dataRow.createCell(2);
                itemCuringCell.setCellValue(p.getITEM_CURING() != null ? p.getITEM_CURING() : "");
                itemCuringCell.setCellStyle(borderStyle);

                Cell patternIdCell = dataRow.createCell(3);
                patternIdCell.setCellValue(p.getPATTERN_ID() != null ? p.getPATTERN_ID().doubleValue() : null);
                patternIdCell.setCellStyle(borderStyle);

                Cell sizeIdCell = dataRow.createCell(4);
                sizeIdCell.setCellValue(p.getSIZE_ID() != null ? p.getSIZE_ID() : "");
                sizeIdCell.setCellStyle(borderStyle);

                Cell productTypeIdCell = dataRow.createCell(5);
                productTypeIdCell.setCellValue(p.getPRODUCT_TYPE_ID() != null ? p.getPRODUCT_TYPE_ID().doubleValue() : null);
                productTypeIdCell.setCellStyle(borderStyle);

                Cell descriptionCell = dataRow.createCell(6);
                descriptionCell.setCellValue(p.getDESCRIPTION() != null ? p.getDESCRIPTION() : "");
                descriptionCell.setCellStyle(borderStyle);

                Cell rimCell = dataRow.createCell(7);
                rimCell.setCellValue(p.getRIM() != null ? p.getRIM().doubleValue() : null);
                rimCell.setCellStyle(borderStyle);

                Cell wibeTubeCell = dataRow.createCell(8);
                wibeTubeCell.setCellValue(p.getWIB_TUBE() != null ? p.getWIB_TUBE() : "");
                wibeTubeCell.setCellStyle(borderStyle);

                Cell itemAssyCell = dataRow.createCell(9);
                itemAssyCell.setCellValue(p.getITEM_ASSY() != null ? p.getITEM_ASSY() : "");
                itemAssyCell.setCellStyle(borderStyle);

                Cell itemExtCell = dataRow.createCell(10);
                itemExtCell.setCellValue(p.getITEM_EXT() != null ? p.getITEM_EXT() : "");
                itemExtCell.setCellStyle(borderStyle);

                Cell extDescriptionCell = dataRow.createCell(11);
                extDescriptionCell.setCellValue(p.getEXT_DESCRIPTION() != null ? p.getEXT_DESCRIPTION() : "");
                extDescriptionCell.setCellStyle(borderStyle);

                Cell qtyPerRakCell = dataRow.createCell(12);
                qtyPerRakCell.setCellValue(p.getQTY_PER_RAK() != null ? p.getQTY_PER_RAK().doubleValue() : null);
                qtyPerRakCell.setCellStyle(borderStyle);

                Cell upperConstantCell = dataRow.createCell(13);
                upperConstantCell.setCellValue(p.getUPPER_CONSTANT() != null ? p.getUPPER_CONSTANT().doubleValue() : null);
                upperConstantCell.setCellStyle(borderStyle);

                Cell lowerConstantCell = dataRow.createCell(14);
                lowerConstantCell.setCellValue(p.getLOWER_CONSTANT() != null ? p.getLOWER_CONSTANT().doubleValue() : null);
                lowerConstantCell.setCellStyle(borderStyle);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Gagal mengekspor data Product");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }


}
