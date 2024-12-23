package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;

import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.servlet.http.HttpServletRequest;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.ItemAssy;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.Pattern;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.ProductType;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.model.TassSize;
import sri.sysint.sri_starter_back.repository.ItemAssyRepo;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.PatternRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.repository.ProductTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;
import sri.sysint.sri_starter_back.service.ProductServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class ProductController {
	private Response response;	

	@Autowired
	private ProductServiceImpl productServiceImpl;
	
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
	@PersistenceContext	
	private EntityManager em;
	
	@GetMapping("/getAllProduct")
	public Response getAllProduct(final HttpServletRequest req) throws ResourceNotFoundException {
		String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
				List<Product> products = new ArrayList<>();
				products = productServiceImpl.getAllProduct();
		
			    response = new Response(
			        new Date(),
			        HttpStatus.OK.value(),
			        null,
			        HttpStatus.OK.getReasonPhrase(),
			        req.getRequestURI(),
			        products
			    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getProductById/{id}")
	public Response getProductById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
	    String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
	        	Optional<Product> product = Optional.of(new Product());
	    	    product = productServiceImpl.getProductById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        product
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveProduct")
	public Response saveProduct(final HttpServletRequest req, @RequestBody Product product) throws ResourceNotFoundException {
	    String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
	        	Product savedProduct = productServiceImpl.saveProduct(product);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedProduct
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateProduct")
	public Response updateProduct(final HttpServletRequest req, @RequestBody Product product) throws ResourceNotFoundException {
	    String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
	        	Product updatedProduct = productServiceImpl.updateProduct(product);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedProduct
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteProduct")
	public Response deleteProduct(final HttpServletRequest req, @RequestBody Product product) throws ResourceNotFoundException {
	    String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
	        	Product deletedProduct= productServiceImpl.deleteProduct(product);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedProduct
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreProduct")
	public Response activateProduct(final HttpServletRequest req, @RequestBody Product product) throws ResourceNotFoundException {
	    String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
	        	Product activatedProduct= productServiceImpl.activateProduct(product);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        activatedProduct
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveProductExcel")
	@Transactional
	public Response saveProductExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
//	    String header = req.getHeader("Authorization");
//
//	    if (header == null || !header.startsWith("Bearer ")) {
//	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
//	    }
//
//	    String token = header.replace("Bearer ", "");
//
//	    try {
//	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
//	            .build()
//	            .verify(token)
//	            .getSubject();
//
//	        if (user != null) {
	            if (file.isEmpty()) {
	                return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "No file uploaded", req.getRequestURI(), null);
	            }

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);

	                List<Product> products = new ArrayList<>();
	                List<String> errorMessages = new ArrayList<>();

	                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);

	                    if (row != null) {
	                        boolean isEmptyRow = true;

	                        for (int j = 0; j < row.getLastCellNum(); j++) {
	                            Cell cell = row.getCell(j);
	                            if (cell != null && cell.getCellType() != CellType.BLANK) {
	                                isEmptyRow = false;
	                                break;
	                            }
	                        }

	                        if (isEmptyRow) {
	                            continue;
	                        }

	                        Product product = new Product();
	                        Cell partnumberCell = row.getCell(1);
	                        Cell itemCuringCell = row.getCell(2);
	                        Cell patternNameCell = row.getCell(3);
	                        Cell sizeCell = row.getCell(4);
	                        Cell productTypeCell = row.getCell(5);
	                        Cell itemAssyCell = row.getCell(9);
	                        Cell descriptionCell = row.getCell(6);
	                        Cell rimCell = row.getCell(7);
	                        Cell wibTubeCell = row.getCell(8);
	                        Cell itemExtCell = row.getCell(10);
	                        Cell extDescriptionCell = row.getCell(11);
	                        Cell qtyPerRakCell = row.getCell(12);
	                        Cell upperConstantCell = row.getCell(13);
	                        Cell lowerConstantCell = row.getCell(14);

	                        
	                        if (partnumberCell == null || partnumberCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 2 (Part Number)");
	                            continue;
	                        }

	                        if (itemCuringCell == null || itemCuringCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Item Curing)");
	                            continue;
	                        }

	                        String itemCuring = itemCuringCell.getStringCellValue();
	                        Optional<ItemCuring> itemCuringOpt = itemCuringRepo.findById(itemCuring);
	                        if (!itemCuringOpt.isPresent()) {
	                            errorMessages.add("Item Curing tidak ditemukan pada Baris " + (i + 1));
	                            continue;
	                        }

	                        if (patternNameCell == null || patternNameCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Pattern Name)");
	                            continue;
	                        }

	                        if (sizeCell == null || sizeCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 5 (Size)");
	                            continue;
	                        }

	                        if (productTypeCell == null || productTypeCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 6 (Product Type)");
	                            continue;
	                        }

	                        if (itemAssyCell == null || itemAssyCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 10 (Item Assy)");
	                            continue;
	                        }

	                        if (descriptionCell == null || descriptionCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 6 (Description)");
	                            continue;
	                        }

	                        if (rimCell == null || rimCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 7 (Rim)");
	                            continue;
	                        }

	                        if (wibTubeCell == null || wibTubeCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 8 (Wib Tube)");
	                            continue;
	                        }

	                        if (itemExtCell == null || itemExtCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 10 (Item Ext)");
	                            continue;
	                        }

	                        if (extDescriptionCell == null || extDescriptionCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 11 (Ext Description)");
	                            continue;
	                        }

	                        if (qtyPerRakCell == null || qtyPerRakCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 12 (Qty Per Rak)");
	                            continue;
	                        }

	                        if (upperConstantCell == null || upperConstantCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 13 (Upper Constant)");
	                            continue;
	                        }

	                        if (lowerConstantCell == null || lowerConstantCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 14 (Lower Constant)");
	                            continue;
	                        }

	                        String patternName = patternNameCell.getStringCellValue();
	                        Optional<Pattern> patternOpt = patternRepo.findByName(patternName);
	                        if (!patternOpt.isPresent()) {
	                            errorMessages.add("Pattern Name tidak ditemukan pada Baris " + (i + 1));
	                            continue;
	                        }

	                        String size = sizeCell.getStringCellValue();
	                        Optional<Size> sizeOpt = sizeRepo.findById(size);
	                        if (!sizeOpt.isPresent()) {
	                            errorMessages.add("Size tidak ditemukan pada Baris " + (i + 1));
	                            continue;
	                        }

	                        String productType = productTypeCell.getStringCellValue();
	                        Optional<ProductType> productTypeOpt = productTypeRepo.findByCategory(productType);
	                        if (!productTypeOpt.isPresent()) {
	                            errorMessages.add("Product Type tidak ditemukan pada Baris " + (i + 1));
	                            continue;
	                        }

	                        String itemAssy = itemAssyCell.getStringCellValue();
	                        Optional<ItemAssy> itemAssyOpt = itemAssyRepo.findById(itemAssy);
	                        if (!itemAssyOpt.isPresent()) {
	                            errorMessages.add("Item Assy tidak ditemukan pada Baris " + (i + 1));
	                            continue;
	                        }

	                        product.setPART_NUMBER(getBigDecimalFromCell(partnumberCell));
	                        product.setITEM_CURING(itemCuring);
	                        product.setPATTERN_ID(patternOpt.get().getPATTERN_ID());
	                        product.setSIZE_ID(size);
	                        product.setPRODUCT_TYPE_ID(productTypeOpt.get().getPRODUCT_TYPE_ID());
	                        product.setDESCRIPTION(getStringFromCell(descriptionCell));
	                        product.setRIM(getBigDecimalFromCell(rimCell));
	                        product.setWIB_TUBE(getStringFromCell(wibTubeCell));
	                        product.setITEM_ASSY(itemAssy);
	                        product.setITEM_EXT(getStringFromCell(itemExtCell));
	                        product.setEXT_DESCRIPTION(getStringFromCell(extDescriptionCell));
	                        product.setQTY_PER_RAK(getBigDecimalFromCell(qtyPerRakCell));
	                        product.setUPPER_CONSTANT(getBigDecimalFromCell(upperConstantCell));
	                        product.setLOWER_CONSTANT(getBigDecimalFromCell(lowerConstantCell));
	                        product.setSTATUS(BigDecimal.valueOf(1));
	                        product.setCREATION_DATE(new Date());
	                        product.setLAST_UPDATE_DATE(new Date());

	                        products.add(product);
	                    }
	                }


	                if (!errorMessages.isEmpty()) {
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
	                }

	                productServiceImpl.deleteAllProduct();
	                for (Product product : products) {
	                    productServiceImpl.saveProduct(product);
	                }

	                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), products);

	            } catch (IOException e) {
	                throw new RuntimeException("Error processing file", e);
	            }
//	        } else {
//	            throw new ResourceNotFoundException("User not found");
//	        }
//	    } catch (IllegalArgumentException e) {
//	        return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, e.getMessage(), req.getRequestURI(), null);
//	    } catch (Exception e) {
//	        throw new ResourceNotFoundException("JWT token is not valid or expired");
//	    }
	}

	private boolean isRowEmpty(Row row) {
	    for (int j = 0; j < row.getLastCellNum(); j++) {
	        Cell cell = row.getCell(j);
	        if (cell != null && cell.getCellType() != CellType.BLANK) {
	            return false;
	        }
	    }
	    return true;
	}

	private BigDecimal getBigDecimalFromCell(Cell cell) {
	    if (cell == null) {
	        return null;
	    }
	    switch (cell.getCellType()) {
	        case STRING:
	            String value = cell.getStringCellValue();
	            try {
	                return new BigDecimal(value);
	            } catch (NumberFormatException e) {
	                System.out.println("Invalid number format in cell: " + value);
	                return null; 
	            }
	        case NUMERIC:
	            return BigDecimal.valueOf(cell.getNumericCellValue());
	        default:
	            return null; 
	    }
	}

	private String getStringFromCell(Cell cell) {
	    if (cell == null) {
	        return null;
	    }
	    return cell.getStringCellValue();
	}
	
    @RequestMapping("/exportProductsExcel")
    public ResponseEntity<InputStreamResource> exportProductsExcel() throws IOException {
        String filename = "EXPORT_MASTER_PRODUCT.xlsx";

        ByteArrayInputStream data = productServiceImpl.exportProductsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    @RequestMapping("/layoutProductsExcel")
    public ResponseEntity<InputStreamResource> layoutProductsExcel() throws IOException {
        String filename = "LAYOUT_MASTER_PRODUCT.xlsx";

        ByteArrayInputStream data = productServiceImpl.layoutProductsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
	
}
