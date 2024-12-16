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
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.MaxCapacity;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.service.MaxCapacityServiceImpl;
import sri.sysint.sri_starter_back.service.PlantServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MaxCapacityController {
		
	private Response response;	

	@Autowired
	private MaxCapacityServiceImpl maxCapacityServiceImpl;
	@Autowired
	private ProductRepo productRepo;
	@Autowired
	private MachineCuringTypeRepo machineCuringTypeRepo;
	@PersistenceContext	
	private EntityManager em;

	//START - GET MAPPING
		@GetMapping("/getAllMaxCapacity")
		public Response getAllMaxCapacity(final HttpServletRequest req) throws ResourceNotFoundException {
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
		        	//function goes here
		        	List<MaxCapacity> maxCapacities = new ArrayList<>();
		    	    maxCapacities = maxCapacityServiceImpl.getAllMaxCapacity();

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        maxCapacities
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@GetMapping("/getMaxCapacityById/{id}")
		public Response getMaxCapacityById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
		        	Optional<MaxCapacity> maxCapacity = Optional.of(new MaxCapacity());
		    	    maxCapacity = maxCapacityServiceImpl.getMaxCapacityById(id);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        maxCapacity
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
	//END - GET MAPPING
	//START - POST MAPPING
		@PostMapping("/saveMaxCapacity")
		public Response saveMaxCapacity(final HttpServletRequest req, @RequestBody MaxCapacity maxCapacity) throws ResourceNotFoundException {
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
		        	MaxCapacity savedMaxCapacity = maxCapacityServiceImpl.saveMaxCapacity(maxCapacity);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        savedMaxCapacity
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/updateMaxCapacity")
		public Response updateMaxCapacity(final HttpServletRequest req, @RequestBody MaxCapacity maxCapacity) throws ResourceNotFoundException {
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
		        	MaxCapacity updatedMaxCapacity = maxCapacityServiceImpl.updateMaxCapacity(maxCapacity);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        updatedMaxCapacity
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/deleteMaxCapacity")
		public Response deleteMaxCapacity(final HttpServletRequest req, @RequestBody MaxCapacity maxCapacity) throws ResourceNotFoundException {
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
		        	MaxCapacity deletedMaxCapacity = maxCapacityServiceImpl.deleteMaxCapacity(maxCapacity);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        deletedMaxCapacity
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/restoreMaxCapacity")
		public Response restoreMaxCapacity(final HttpServletRequest req, @RequestBody MaxCapacity maxCapacity) throws ResourceNotFoundException {
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
		            MaxCapacity restoredMaxCapacity = maxCapacityServiceImpl.restoreMaxCapacity(maxCapacity);

		            Response response = new Response(
		                new Date(),
		                HttpStatus.OK.value(),
		                null,
		                HttpStatus.OK.getReasonPhrase(),
		                req.getRequestURI(),
		                restoredMaxCapacity
		            );
		            return response;
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }
		}

		
		@PostMapping("/saveMaxCapacitiesExcel")
		@Transactional
		public Response saveMaxCapacitiesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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
		            if (file.isEmpty()) {
		                return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "No file uploaded", req.getRequestURI(), null);
		            }

		            try (InputStream inputStream = file.getInputStream()) {
		                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		                XSSFSheet sheet = workbook.getSheetAt(0);

		                List<MaxCapacity> maxCapacities = new ArrayList<>();
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

		                        MaxCapacity maxCapacity = new MaxCapacity();
		                        Cell productIdCell = row.getCell(2);
		                        Cell machineCurringTypeIDCell = row.getCell(3);
		                        Cell cycleTimeCell = row.getCell(4);
		                        Cell capacityShift1Cell = row.getCell(5);
		                        Cell capacityShift2Cell = row.getCell(6);
		                        Cell capacityShift3Cell = row.getCell(7);

		                        if (productIdCell == null || productIdCell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 3 (Product ID)");
		                            continue;
		                        }

		                        if (machineCurringTypeIDCell == null || machineCurringTypeIDCell.getCellType() != CellType.STRING) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 4 (Machine Curring Type ID)");
		                            continue;
		                        }

		                        if (cycleTimeCell == null || cycleTimeCell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 5 (Cycle Time)");
		                            continue;
		                        }

		                        if (capacityShift1Cell == null || capacityShift1Cell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 6 (Capacity Shift 1)");
		                            continue;
		                        }

		                        if (capacityShift2Cell == null || capacityShift2Cell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 7 (Capacity Shift 2)");
		                            continue;
		                        }

		                        if (capacityShift3Cell == null || capacityShift3Cell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 8 (Capacity Shift 3)");
		                            continue;
		                        }

		                        Optional<Product> productOpt = productRepo.findById(BigDecimal.valueOf(productIdCell.getNumericCellValue()));
		                        Optional<MachineCuringType> machineCuringTypeOpt = machineCuringTypeRepo.findById(machineCurringTypeIDCell.getStringCellValue());

		                        if (productOpt.isEmpty()) {
		                            errorMessages.add("Data Tidak Valid, Data Product pada Baris " + (i + 1) + " Tidak Ditemukan");
		                            continue;
		                        }

		                        if (machineCuringTypeOpt.isEmpty()) {
		                            errorMessages.add("Data Tidak Valid, Data Machine Curring Type pada Baris " + (i + 1) + " Tidak Ditemukan");
		                            continue;
		                        }

		                        maxCapacity.setMAX_CAP_ID(maxCapacityServiceImpl.getNewId());
		                        maxCapacity.setPRODUCT_ID(BigDecimal.valueOf(productIdCell.getNumericCellValue()));
		                        maxCapacity.setMACHINECURINGTYPE_ID(machineCurringTypeIDCell.getStringCellValue());
		                        maxCapacity.setCYCLE_TIME(BigDecimal.valueOf(cycleTimeCell.getNumericCellValue()));
		                        maxCapacity.setCAPACITY_SHIFT_1(BigDecimal.valueOf(capacityShift1Cell.getNumericCellValue()));
		                        maxCapacity.setCAPACITY_SHIFT_2(BigDecimal.valueOf(capacityShift2Cell.getNumericCellValue()));
		                        maxCapacity.setCAPACITY_SHIFT_3(BigDecimal.valueOf(capacityShift3Cell.getNumericCellValue()));
		                        maxCapacity.setSTATUS(BigDecimal.valueOf(1));
		                        maxCapacity.setCREATION_DATE(new Date());
		                        maxCapacity.setLAST_UPDATE_DATE(new Date());

		                        maxCapacities.add(maxCapacity);
		                    }
		                }

		                if (!errorMessages.isEmpty()) {
		                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
		                }

		                maxCapacityServiceImpl.deleteAllMaxCapacity();
		                for (MaxCapacity maxCapacity : maxCapacities) {
		                    maxCapacityServiceImpl.saveMaxCapacity(maxCapacity);
		                }

		                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), maxCapacities);

		            } catch (IOException e) {
		                throw new RuntimeException("Error processing file", e);
		            }
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (IllegalArgumentException e) {
		        return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, e.getMessage(), req.getRequestURI(), null);
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }
		}
		
	    @RequestMapping("/exportMaxCapacityExcel")
	    public ResponseEntity<InputStreamResource> exportMaxCapacityExcel() throws IOException {
	        String filename = "EXPORT_MASTER_MAX_CAPACITY.xlsx"; 

	        ByteArrayInputStream data = maxCapacityServiceImpl.exportMaxCapacitysExcel(); 
	        InputStreamResource file = new InputStreamResource(data);

	        return ResponseEntity.ok()
	                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename) 
	                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) 
	                .body(file); 
	    }
	    
	    @RequestMapping("/layoutMaxCapacityExcel")
	    public ResponseEntity<InputStreamResource> layoutMaxCapacityExcel() throws IOException {
	        String filename = "LAYOUT_MASTER_MAX_CAPACITY.xlsx"; 

	        ByteArrayInputStream data = maxCapacityServiceImpl.layoutMaxCapacitysExcel(); 
	        InputStreamResource file = new InputStreamResource(data);

	        return ResponseEntity.ok()
	                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename) 
	                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) 
	                .body(file); 
	    }

//END - POST MAPPING
//START - PUT MAPPING
//END - PUT MAPPING
//START - DELETE MAPPING
//END - DELETE MAPPING
//START - PROCEDURE
//END - PROCEDURE
}


