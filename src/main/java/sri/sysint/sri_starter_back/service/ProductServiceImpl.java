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

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.ItemAssy;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.Pattern;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.ProductType;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.ItemAssyRepo;
import sri.sysint.sri_starter_back.repository.PatternRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.repository.ProductTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;

@Service
@Transactional
public class ProductServiceImpl {
	
	@Autowired
    private ProductRepo productRepo;
	
	@Autowired
    private PatternRepo patternRepo;
	
	@Autowired
    private ItemCuringRepo itemCuringRepo;
	
	@Autowired
    private ItemAssyRepo itemAssyRepo;
	
	@Autowired
    private SizeRepo sizeRepo;
	
	@Autowired
    private ProductTypeRepo productTypeRepo;
	
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
            "PATTERN_NAME",
            "SIZE",
            "CATEGORY",
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
            List<String> itemCuringList = itemCuringRepo.findItemCuringActive().stream().map(ItemCuring::getITEM_CURING).collect(Collectors.toList());
            List<String> patternList = patternRepo.findPatternActive().stream().map(Pattern::getPATTERN_NAME).collect(Collectors.toList());
            List<String> sizeList = sizeRepo.findSizeActive().stream().map(Size::getSIZE_ID).collect(Collectors.toList());
            List<String> productTypeList = productTypeRepo.findProductTypeActive().stream().map(ProductType::getCATEGORY).collect(Collectors.toList());
            List<String> itemAssyList = itemAssyRepo.findItemAssyActive().stream().map(ItemAssy::getITEM_ASSY).collect(Collectors.toList());

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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_DATA");
            int hiddenRowIndex = 0;

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, itemCuringList, hiddenRowIndex);
            Name itemCuringName = workbook.createName();
            itemCuringName.setNameName("ItemCuringList");
            itemCuringName.setRefersToFormula("HIDDEN_DATA!$A$1:$A$" + itemCuringList.size());

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, patternList, hiddenRowIndex);
            Name patternName = workbook.createName();
            patternName.setNameName("PatternList");
            patternName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - patternList.size() + 1) + ":$A$" + hiddenRowIndex);

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, sizeList, hiddenRowIndex);
            Name sizeName = workbook.createName();
            sizeName.setNameName("SizeList");
            sizeName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - sizeList.size() + 1) + ":$A$" + hiddenRowIndex);

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, productTypeList, hiddenRowIndex);
            Name productTypeName = workbook.createName();
            productTypeName.setNameName("ProductTypeList");
            productTypeName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - productTypeList.size() + 1) + ":$A$" + hiddenRowIndex);

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, itemAssyList, hiddenRowIndex);
            Name itemAssyName = workbook.createName();
            itemAssyName.setNameName("ItemAssyList");
            itemAssyName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - itemAssyList.size() + 1) + ":$A$" + hiddenRowIndex);

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

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

                Cell patternNameCell = dataRow.createCell(3);
                patternNameCell.setCellValue(p.getPATTERN_ID() != null ? patternRepo.findById(p.getPATTERN_ID()).map(Pattern::getPATTERN_NAME).orElse("") : "");
                patternNameCell.setCellStyle(borderStyle);

                Cell sizeIdCell = dataRow.createCell(4);
                sizeIdCell.setCellValue(p.getSIZE_ID() != null ? p.getSIZE_ID() : "");
                sizeIdCell.setCellStyle(borderStyle);

                Cell categoryCell = dataRow.createCell(5);
                categoryCell.setCellValue(p.getPRODUCT_TYPE_ID() != null ? productTypeRepo.findById(p.getPRODUCT_TYPE_ID()).map(ProductType::getCATEGORY).orElse("") : "");
                categoryCell.setCellStyle(borderStyle);

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

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            addValidation(sheet, "ItemCuringList", 2);
            addValidation(sheet, "PatternList", 3);
            addValidation(sheet, "SizeList", 4);
            addValidation(sheet, "ProductTypeList", 5);
            addValidation(sheet, "ItemAssyList", 9);

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

    private int populateHiddenSheet(Sheet hiddenSheet, List<String> dataList, int startIndex) {
        for (int i = 0; i < dataList.size(); i++) {
            Row row = hiddenSheet.createRow(startIndex + i);
            Cell cell = row.createCell(0);
            cell.setCellValue(dataList.get(i));
        }
        return startIndex + dataList.size();
    }

    private void addValidation(Sheet sheet, String listName, int column) {
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = validationHelper.createFormulaListConstraint(listName);
        CellRangeAddressList addressList = new CellRangeAddressList(1, 1000, column, column);
        DataValidation validation = validationHelper.createValidation(constraint, addressList);
        validation.setSuppressDropDownArrow(true);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    public ByteArrayInputStream layoutProductsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel();
        return byteArrayInputStream;
    }
    
    private ByteArrayInputStream layoutToExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "PART_NUMBER",
            "ITEM_CURING",
            "PATTERN_NAME",
            "SIZE",
            "CATEGORY",
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
            List<String> itemCuringList = itemCuringRepo.findItemCuringActive().stream().map(ItemCuring::getITEM_CURING).collect(Collectors.toList());
            List<String> patternList = patternRepo.findPatternActive().stream().map(Pattern::getPATTERN_NAME).collect(Collectors.toList());
            List<String> sizeList = sizeRepo.findSizeActive().stream().map(Size::getSIZE_ID).collect(Collectors.toList());
            List<String> productTypeList = productTypeRepo.findProductTypeActive().stream().map(ProductType::getCATEGORY).collect(Collectors.toList());
            List<String> itemAssyList = itemAssyRepo.findItemAssyActive().stream().map(ItemAssy::getITEM_ASSY).collect(Collectors.toList());

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

            Sheet hiddenSheet = workbook.createSheet("HIDDEN_DATA");
            int hiddenRowIndex = 0;

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, itemCuringList, hiddenRowIndex);
            Name itemCuringName = workbook.createName();
            itemCuringName.setNameName("ItemCuringList");
            itemCuringName.setRefersToFormula("HIDDEN_DATA!$A$1:$A$" + itemCuringList.size());

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, patternList, hiddenRowIndex);
            Name patternName = workbook.createName();
            patternName.setNameName("PatternList");
            patternName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - patternList.size() + 1) + ":$A$" + hiddenRowIndex);

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, sizeList, hiddenRowIndex);
            Name sizeName = workbook.createName();
            sizeName.setNameName("SizeList");
            sizeName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - sizeList.size() + 1) + ":$A$" + hiddenRowIndex);

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, productTypeList, hiddenRowIndex);
            Name productTypeName = workbook.createName();
            productTypeName.setNameName("ProductTypeList");
            productTypeName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - productTypeList.size() + 1) + ":$A$" + hiddenRowIndex);

            hiddenRowIndex = populateHiddenSheet(hiddenSheet, itemAssyList, hiddenRowIndex);
            Name itemAssyName = workbook.createName();
            itemAssyName.setNameName("ItemAssyList");
            itemAssyName.setRefersToFormula("HIDDEN_DATA!$A$" + (hiddenRowIndex - itemAssyList.size() + 1) + ":$A$" + hiddenRowIndex);

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);

            int rowIndex = 1;
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }
            
            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }
            addValidation(sheet, "ItemCuringList", 2);
            addValidation(sheet, "PatternList", 3);
            addValidation(sheet, "SizeList", 4);
            addValidation(sheet, "ProductTypeList", 5);
            addValidation(sheet, "ItemAssyList", 9);

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
