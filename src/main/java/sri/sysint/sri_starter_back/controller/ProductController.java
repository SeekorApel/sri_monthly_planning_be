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
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.TassSize;
import sri.sysint.sri_starter_back.service.ProductServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class ProductController {
	private Response response;	

	@Autowired
	private ProductServiceImpl productServiceImpl;
	
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
	            
	            productServiceImpl.deleteAllProduct();

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);
	                
	                List<Product> products = new ArrayList<>();
	                
	                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);
	                    if (row == null || isRowEmpty(row)) {
	                        continue; // Skip empty rows
	                    }

	                    Product product = new Product();
	                    product.setPART_NUMBER(getBigDecimalFromCell(row.getCell(0))); // Part number
	                    product.setITEM_CURING(getStringFromCell(row.getCell(1)));
	                    product.setPATTERN_ID(getBigDecimalFromCell(row.getCell(2)));
	                    product.setSIZE_ID(getStringFromCell(row.getCell(3)));
	                    product.setPRODUCT_TYPE_ID(getBigDecimalFromCell(row.getCell(4)));
	                    product.setDESCRIPTION(getStringFromCell(row.getCell(5)));
	                    product.setRIM(getBigDecimalFromCell(row.getCell(6)));
	                    product.setWIB_TUBE(getStringFromCell(row.getCell(7)));
	                    product.setITEM_ASSY(getStringFromCell(row.getCell(8)));
	                    product.setITEM_EXT(getStringFromCell(row.getCell(9)));
	                    product.setEXT_DESCRIPTION(getStringFromCell(row.getCell(10)));
	                    product.setQTY_PER_RAK(getBigDecimalFromCell(row.getCell(11)));
	                    product.setUPPER_CONSTANT(getBigDecimalFromCell(row.getCell(12)));
	                    product.setLOWER_CONSTANT(getBigDecimalFromCell(row.getCell(13)));
	                    product.setSTATUS(BigDecimal.valueOf(1));
	                    product.setCREATION_DATE(new Date());
	                    product.setLAST_UPDATE_DATE(new Date());

	                    productServiceImpl.saveProduct(product);
	                    products.add(product);
	                }

	                response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), products);
	                
	            } catch (IOException e) {
	                response = new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file", req.getRequestURI(), null);
	            }

//	        } else {
//	            throw new ResourceNotFoundException("User not found");
//	        }
//	    } catch (Exception e) {
//	        throw new ResourceNotFoundException("JWT token is not valid or expired");
//	    }

	    return response;
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
	                return null; // Handle it as needed
	            }
	        case NUMERIC:
	            return BigDecimal.valueOf(cell.getNumericCellValue());
	        default:
	            return null; // Handle it as needed
	    }
	}

	private String getStringFromCell(Cell cell) {
	    if (cell == null) {
	        return null;
	    }
	    
	    switch (cell.getCellType()) {
	        case STRING:
	            return cell.getStringCellValue();
	        case NUMERIC:
	            return String.valueOf(cell.getNumericCellValue());
	        case BOOLEAN:
	            return String.valueOf(cell.getBooleanCellValue());
	        default:
	            return null; // Handle other cell types as needed
	    }
	}
	
    @RequestMapping("/exportProductsExcel")
    public ResponseEntity<InputStreamResource> exportProductsExcel() throws IOException {
        String filename = "MASTER_PRODUCT.xlsx";

        ByteArrayInputStream data = productServiceImpl.exportProductsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
	
}
