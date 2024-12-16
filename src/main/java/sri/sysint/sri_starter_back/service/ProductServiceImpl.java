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
import sri.sysint.sri_starter_back.repository.PatternRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.repository.ProductTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.ItemAssyRepo;

@Service
@Transactional
public class ProductServiceImpl {
	
	@Autowired
    private ProductRepo productRepo;
	
	@Autowired
    private PatternRepo patternRepo;
	
	@Autowired
    private ProductTypeRepo productTypeRepo;

    @Autowired
    private SizeRepo sizeRepo;
	
	@Autowired
    private ItemCuringRepo itemCuringRepo;

    @Autowired
    private ItemAssyRepo itemAssyRepo;
	
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
                        List<ItemCuring> activeItemCurings = itemCuringRepo.findItemCuringActive();
            List<Pattern> activePatterns = patternRepo.findPatternActive();
            List<Size> activeSizes = sizeRepo.findSizeActive();
            List<ProductType> activeProductTypes = productTypeRepo.findProductTypeActive();
            List<ItemAssy> activeItemAssys = itemAssyRepo.findItemAssyActive();

            List<String> itemCuringID = activeItemCurings.stream()
                .map(ItemCuring::getITEM_CURING)
                .collect(Collectors.toList());

            List<String> patternName = activePatterns.stream()
                .map(Pattern::getPATTERN_NAME)
                .collect(Collectors.toList());

            List<String> sizeID = activeSizes.stream()
                .map(Size::getSIZE_ID)
                .collect(Collectors.toList());

            List<String> productTypeCategory = activeProductTypes.stream()
                .map(ProductType::getCATEGORY)
                .collect(Collectors.toList());

            List<String> itemAssyID = activeItemAssys.stream()
                .map(ItemAssy::getITEM_ASSY)
                .collect(Collectors.toList());

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
             Sheet hiddenSheetCuring = workbook.createSheet("HIDDEN_ITEMCURINGS");
            for (int i = 0; i < itemCuringID.size(); i++) {
                Row row = hiddenSheetCuring.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(itemCuringID.get(i));
            }

            Name namedRangeCuring = workbook.createName();
            namedRangeCuring.setNameName("ItemCuringID");
            namedRangeCuring.setRefersToFormula("HIDDEN_ITEMCURINGS!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetCuring), true);

            Sheet hiddenSheetPattern = workbook.createSheet("HIDDEN_PATTERNS");
            for (int i = 0; i < patternName.size(); i++) {
                Row row = hiddenSheetPattern.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(patternName.get(i));
            }

            Name namedRangePattern = workbook.createName();
            namedRangePattern.setNameName("patternName");
            namedRangePattern.setRefersToFormula("HIDDEN_PATTERNS!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetPattern), true);

            Sheet hiddenSheetSize = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < sizeID.size(); i++) {
                Row row = hiddenSheetSize.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(sizeID.get(i));
            }

            Name namedRangeSize = workbook.createName();
            namedRangeSize.setNameName("sizeID");
            namedRangeSize.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetSize), true);

            Sheet hiddenSheetProductType = workbook.createSheet("HIDDEN_PRODUCTTYPES");
            for (int i = 0; i < productTypeCategory.size(); i++) {
                Row row = hiddenSheetProductType.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(productTypeCategory.get(i));
            }

            Name namedRangeProductType = workbook.createName();
            namedRangeProductType.setNameName("productTypeCategory");
            namedRangeProductType.setRefersToFormula("HIDDEN_PRODUCTTYPES!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetProductType), true);

            Sheet hiddenSheetAssy = workbook.createSheet("HIDDEN_ITEMASSYS");
            for (int i = 0; i < itemAssyID.size(); i++) {
                Row row = hiddenSheetAssy.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(itemAssyID.get(i));
            }

            Name namedRangeAssy = workbook.createName();
            namedRangeAssy.setNameName("ItemAssyID");
            namedRangeAssy.setRefersToFormula("HIDDEN_ITEMASSYS!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetAssy), true);

            int rowIndex = 1;
            for (Product p : products) {
                Row dataRow = sheet.createRow(rowIndex++);

                // Nomor
                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(rowIndex - 1);
                nomorCell.setCellStyle(borderStyle);

                // PART_NUMBER
                Cell partNumberCell = dataRow.createCell(1);
                partNumberCell.setCellValue(p.getPART_NUMBER().doubleValue());
                partNumberCell.setCellStyle(borderStyle);

                // ITEM_CURING
                Cell itemCuringCell = dataRow.createCell(2);
                itemCuringCell.setCellValue(p.getITEM_CURING() != null ? p.getITEM_CURING() : "");
                itemCuringCell.setCellStyle(borderStyle);

                // PATTERN_NAME
                Cell patternNameCell = dataRow.createCell(3);
                if (p.getPATTERN_ID() != null) {
                    Optional<Pattern> pattern = patternRepo.findById(p.getPATTERN_ID());
                    patternNameCell.setCellValue(pattern.isPresent() ? pattern.get().getPATTERN_NAME() : "");
                }
                patternNameCell.setCellStyle(borderStyle);

                // SIZE_ID (still SIZE_ID)
                Cell sizeIdCell = dataRow.createCell(4);
                sizeIdCell.setCellValue(p.getSIZE_ID() != null ? p.getSIZE_ID() : "");
                sizeIdCell.setCellStyle(borderStyle);

                // CATEGORY (PRODUCT_TYPE_ID -> CATEGORY)
                Cell categoryCell = dataRow.createCell(5);
                if (p.getPRODUCT_TYPE_ID() != null) {
                    Optional<ProductType> productType = productTypeRepo.findById(p.getPRODUCT_TYPE_ID());
                    categoryCell.setCellValue(productType.isPresent() ? productType.get().getCATEGORY() : "");
                }
                categoryCell.setCellStyle(borderStyle);

                // DESCRIPTION
                Cell descriptionCell = dataRow.createCell(6);
                descriptionCell.setCellValue(p.getDESCRIPTION() != null ? p.getDESCRIPTION() : "");
                descriptionCell.setCellStyle(borderStyle);

                // RIM
                Cell rimCell = dataRow.createCell(7);
                rimCell.setCellValue(p.getRIM() != null ? p.getRIM().doubleValue() : null);
                rimCell.setCellStyle(borderStyle);

                // WIB_TUBE
                Cell wibeTubeCell = dataRow.createCell(8);
                wibeTubeCell.setCellValue(p.getWIB_TUBE() != null ? p.getWIB_TUBE() : "");
                wibeTubeCell.setCellStyle(borderStyle);

                // ITEM_ASSY
                Cell itemAssyCell = dataRow.createCell(9);
                itemAssyCell.setCellValue(p.getITEM_ASSY() != null ? p.getITEM_ASSY() : "");
                itemAssyCell.setCellStyle(borderStyle);

                // ITEM_EXT
                Cell itemExtCell = dataRow.createCell(10);
                itemExtCell.setCellValue(p.getITEM_EXT() != null ? p.getITEM_EXT() : "");
                itemExtCell.setCellStyle(borderStyle);

                // EXT_DESCRIPTION
                Cell extDescriptionCell = dataRow.createCell(11);
                extDescriptionCell.setCellValue(p.getEXT_DESCRIPTION() != null ? p.getEXT_DESCRIPTION() : "");
                extDescriptionCell.setCellStyle(borderStyle);

                // QTY_PER_RAK
                Cell qtyPerRakCell = dataRow.createCell(12);
                qtyPerRakCell.setCellValue(p.getQTY_PER_RAK() != null ? p.getQTY_PER_RAK().doubleValue() : null);
                qtyPerRakCell.setCellStyle(borderStyle);

                // UPPER_CONSTANT
                Cell upperConstantCell = dataRow.createCell(13);
                upperConstantCell.setCellValue(p.getUPPER_CONSTANT() != null ? p.getUPPER_CONSTANT().doubleValue() : null);
                upperConstantCell.setCellStyle(borderStyle);

                // LOWER_CONSTANT
                Cell lowerConstantCell = dataRow.createCell(14);
                lowerConstantCell.setCellValue(p.getLOWER_CONSTANT() != null ? p.getLOWER_CONSTANT().doubleValue() : null);
                lowerConstantCell.setCellStyle(borderStyle);
            }

                        DataValidationHelper validationHelper1 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint1 = validationHelper1.createFormulaListConstraint("ItemCuringID");
            CellRangeAddressList addressList1 = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation1 = validationHelper1.createValidation(constraint1, addressList1);
            validation1.setSuppressDropDownArrow(true);
            validation1.setShowErrorBox(true);
            sheet.addValidationData(validation1);

            DataValidationHelper validationHelper2 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint2 = validationHelper2.createFormulaListConstraint("patternName");
            CellRangeAddressList addressList2 = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation validation2 = validationHelper2.createValidation(constraint2, addressList2);
            validation2.setSuppressDropDownArrow(true);
            validation2.setShowErrorBox(true);
            sheet.addValidationData(validation2);

            DataValidationHelper validationHelper3 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint3 = validationHelper3.createFormulaListConstraint("sizeID");
            CellRangeAddressList addressList3 = new CellRangeAddressList(1, 1000, 4, 4);
            DataValidation validation3 = validationHelper3.createValidation(constraint3, addressList3);
            validation3.setSuppressDropDownArrow(true);
            validation3.setShowErrorBox(true);
            sheet.addValidationData(validation3);

            DataValidationHelper validationHelper4 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint4 = validationHelper4.createFormulaListConstraint("productTypeCategory");
            CellRangeAddressList addressList4 = new CellRangeAddressList(1, 1000, 5, 5);
            DataValidation validation4 = validationHelper4.createValidation(constraint4, addressList4);
            validation4.setSuppressDropDownArrow(true);
            validation4.setShowErrorBox(true);
            sheet.addValidationData(validation4);

            DataValidationHelper validationHelper5 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint5 = validationHelper5.createFormulaListConstraint("itemAssyID");
            CellRangeAddressList addressList5 = new CellRangeAddressList(1, 1000, 9, 9);
            DataValidation validation5 = validationHelper5.createValidation(constraint5, addressList5);
            validation5.setSuppressDropDownArrow(true);
            validation5.setShowErrorBox(true);
            sheet.addValidationData(validation5);

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

    public ByteArrayInputStream layoutProductsExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = layoutToExcel( );
        return byteArrayInputStream;
    }
    
    public ByteArrayInputStream layoutToExcel() throws IOException {
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
            List<ItemCuring> activeItemCurings = itemCuringRepo.findItemCuringActive();
            List<Pattern> activePatterns = patternRepo.findPatternActive();
            List<Size> activeSizes = sizeRepo.findSizeActive();
            List<ProductType> activeProductTypes = productTypeRepo.findProductTypeActive();
            List<ItemAssy> activeItemAssys = itemAssyRepo.findItemAssyActive();

            List<String> itemCuringID = activeItemCurings.stream()
                .map(ItemCuring::getITEM_CURING)
                .collect(Collectors.toList());

            List<String> patternName = activePatterns.stream()
                .map(Pattern::getPATTERN_NAME)
                .collect(Collectors.toList());

            List<String> sizeID = activeSizes.stream()
                .map(Size::getSIZE_ID)
                .collect(Collectors.toList());

            List<String> productTypeCategory = activeProductTypes.stream()
                .map(ProductType::getCATEGORY)
                .collect(Collectors.toList());

            List<String> itemAssyID = activeItemAssys.stream()
                .map(ItemAssy::getITEM_ASSY)
                .collect(Collectors.toList());

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
            
            for (int i = 1; i <= 5; i++) {
                Row dataRow = sheet.createRow(i);
                for (int j = 0; j < header.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellStyle(borderStyle);
                }
            }

            Sheet hiddenSheetCuring = workbook.createSheet("HIDDEN_ITEMCURINGS");
            for (int i = 0; i < itemCuringID.size(); i++) {
                Row row = hiddenSheetCuring.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(itemCuringID.get(i));
            }

            Name namedRangeCuring = workbook.createName();
            namedRangeCuring.setNameName("ItemCuringID");
            namedRangeCuring.setRefersToFormula("HIDDEN_ITEMCURINGS!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetCuring), true);

            Sheet hiddenSheetPattern = workbook.createSheet("HIDDEN_PATTERNS");
            for (int i = 0; i < patternName.size(); i++) {
                Row row = hiddenSheetPattern.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(patternName.get(i));
            }

            Name namedRangePattern = workbook.createName();
            namedRangePattern.setNameName("patternName");
            namedRangePattern.setRefersToFormula("HIDDEN_PATTERNS!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetPattern), true);

            Sheet hiddenSheetSize = workbook.createSheet("HIDDEN_SIZES");
            for (int i = 0; i < sizeID.size(); i++) {
                Row row = hiddenSheetSize.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(sizeID.get(i));
            }

            Name namedRangeSize = workbook.createName();
            namedRangeSize.setNameName("sizeID");
            namedRangeSize.setRefersToFormula("HIDDEN_SIZES!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetSize), true);

            Sheet hiddenSheetProductType = workbook.createSheet("HIDDEN_PRODUCTTYPES");
            for (int i = 0; i < productTypeCategory.size(); i++) {
                Row row = hiddenSheetProductType.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(productTypeCategory.get(i));
            }

            Name namedRangeProductType = workbook.createName();
            namedRangeProductType.setNameName("productTypeCategory");
            namedRangeProductType.setRefersToFormula("HIDDEN_PRODUCTTYPES!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetProductType), true);

            Sheet hiddenSheetAssy = workbook.createSheet("HIDDEN_ITEMASSYS");
            for (int i = 0; i < itemAssyID.size(); i++) {
                Row row = hiddenSheetAssy.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(itemAssyID.get(i));
            }

            Name namedRangeAssy = workbook.createName();
            namedRangeAssy.setNameName("ItemAssyID");
            namedRangeAssy.setRefersToFormula("HIDDEN_ITEMASSYS!$A$1:$A$" + itemCuringID.size());

            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheetAssy), true);

            int rowIndex = 1;
            int nomor = 1;
            
            DataValidationHelper validationHelper1 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint1 = validationHelper1.createFormulaListConstraint("ItemCuringID");
            CellRangeAddressList addressList1 = new CellRangeAddressList(1, 1000, 2, 2);
            DataValidation validation1 = validationHelper1.createValidation(constraint1, addressList1);
            validation1.setSuppressDropDownArrow(true);
            validation1.setShowErrorBox(true);
            sheet.addValidationData(validation1);

            DataValidationHelper validationHelper2 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint2 = validationHelper2.createFormulaListConstraint("patternName");
            CellRangeAddressList addressList2 = new CellRangeAddressList(1, 1000, 3, 3);
            DataValidation validation2 = validationHelper2.createValidation(constraint2, addressList2);
            validation2.setSuppressDropDownArrow(true);
            validation2.setShowErrorBox(true);
            sheet.addValidationData(validation2);

            DataValidationHelper validationHelper3 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint3 = validationHelper3.createFormulaListConstraint("sizeID");
            CellRangeAddressList addressList3 = new CellRangeAddressList(1, 1000, 4, 4);
            DataValidation validation3 = validationHelper3.createValidation(constraint3, addressList3);
            validation3.setSuppressDropDownArrow(true);
            validation3.setShowErrorBox(true);
            sheet.addValidationData(validation3);

            DataValidationHelper validationHelper4 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint4 = validationHelper4.createFormulaListConstraint("productTypeCategory");
            CellRangeAddressList addressList4 = new CellRangeAddressList(1, 1000, 5, 5);
            DataValidation validation4 = validationHelper4.createValidation(constraint4, addressList4);
            validation4.setSuppressDropDownArrow(true);
            validation4.setShowErrorBox(true);
            sheet.addValidationData(validation4);

            DataValidationHelper validationHelper5 = sheet.getDataValidationHelper();
            DataValidationConstraint constraint5 = validationHelper5.createFormulaListConstraint("itemAssyID");
            CellRangeAddressList addressList5 = new CellRangeAddressList(1, 1000, 9, 9);
            DataValidation validation5 = validationHelper5.createValidation(constraint5, addressList5);
            validation5.setSuppressDropDownArrow(true);
            validation5.setShowErrorBox(true);
            sheet.addValidationData(validation5);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export Quadrant data");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }

}
