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
import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;
import sri.sysint.sri_starter_back.service.MachineCuringServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineCuringController {
		
	private Response response;	

	@Autowired
	private MachineCuringServiceImpl machineCuringServiceImpl;
    @Autowired
    private MachineCuringTypeRepo machineCuringTypeRepo;
    @Autowired
    private BuildingRepo buildingRepo;
	@PersistenceContext	
	private EntityManager em;
	
	//START - GET MAPPING
		@GetMapping("/getAllMachineCuring")
		public Response getAllMachineCuring(final HttpServletRequest req) throws ResourceNotFoundException {
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
		        	List<MachineCuring> machineCurings = new ArrayList<>();
		        	machineCurings = machineCuringServiceImpl.getAllMachineCuring();

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        machineCurings
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@GetMapping("/getMachineCuringById/{id}")
		public Response getMachineCuringById(final HttpServletRequest req, @PathVariable String id) throws ResourceNotFoundException {
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
		        	Optional<MachineCuring> machineCuring = Optional.of(new MachineCuring());
		        	machineCuring = machineCuringServiceImpl.getMachineCuringById(id);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        machineCuring
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
		@PostMapping("/saveMachineCuring")
		public Response saveMachineCuring(final HttpServletRequest req, @RequestBody MachineCuring machineCuring) throws ResourceNotFoundException {
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
		        	MachineCuring savedMachineCuring = machineCuringServiceImpl.saveMachineCuring(machineCuring);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        savedMachineCuring
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/updateMachineCuring")
		public Response updateMachineCuring(final HttpServletRequest req, @RequestBody MachineCuring machineCuring) throws ResourceNotFoundException {
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
		        	MachineCuring updatedMachineCuring = machineCuringServiceImpl.updateMachineCuring(machineCuring);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        updatedMachineCuring
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/deleteMachineCuring")
		public Response deleteteMachineCuring(final HttpServletRequest req, @RequestBody MachineCuring machineCuring) throws ResourceNotFoundException {
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
		        	MachineCuring deletedMachineCuring = machineCuringServiceImpl.deleteMachineCuring(machineCuring);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        deletedMachineCuring
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/restoreMachineCuring")
		public Response restoreMachineCuring(final HttpServletRequest req, @RequestBody MachineCuring machineCuring) throws ResourceNotFoundException {
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
		            MachineCuring restoredMachineCuring = machineCuringServiceImpl.restoreMachineCuring(machineCuring);

		            response = new Response(
		                new Date(),
		                HttpStatus.OK.value(),
		                null,
		                HttpStatus.OK.getReasonPhrase(),
		                req.getRequestURI(),
		                restoredMachineCuring
		            );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/saveMachineCuringExcel")
		@Transactional
		public Response saveMachineCuringExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

		                List<MachineCuring> machineCurings = new ArrayList<>();
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

		                        Cell workCenterTextCell = row.getCell(1);
		                        Cell buildingNameCell = row.getCell(2);
		                        Cell cavityCell = row.getCell(3);
		                        Cell machinetype = row.getCell(4);
		                        Cell statususage = row.getCell(5);

		                        if (workCenterTextCell == null || workCenterTextCell.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 2 (Work Center Text)");
		                            continue;
		                        }

		                        if (buildingNameCell == null || buildingNameCell.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Building Name)");
		                            continue;
		                        }

		                        if (cavityCell == null || cavityCell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Bukan Numerik pada Baris " + (i + 1) + " Kolom 4 (Cavity)");
		                            continue;
		                        }

		                        if (machinetype == null || machinetype.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 5 (Machine Type)");
		                            continue;
		                        }

		                        if (statususage == null || statususage.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Bukan Numerik pada Baris " + (i + 1) + " Kolom 6 (Status Usage)");
		                            continue;
		                        }

		                        String buildingName = buildingNameCell.getStringCellValue();
		                        Optional<Building> buildingOptional = buildingRepo.findByName(buildingName);

		                        if (buildingOptional.isEmpty()) {
		                            errorMessages.add("Data Tidak Valid, Building Name pada Baris " + (i + 1) + " Tidak Ditemukan");
		                            continue;
		                        }

		                        String machineTypeValue = machinetype.getStringCellValue();
		                        Optional<MachineCuringType> productTypeOptional = machineCuringTypeRepo.findById(machineTypeValue);

		                        if (productTypeOptional.isEmpty()) {
		                            errorMessages.add("Data Tidak Valid, Machine Type pada Baris " + (i + 1) + " Tidak Ditemukan");
		                            continue;
		                        }

		                        MachineCuring machineCuring = new MachineCuring();

		                        machineCuring.setWORK_CENTER_TEXT(workCenterTextCell.getStringCellValue());
		                        machineCuring.setBUILDING_ID(buildingOptional.get().getBUILDING_ID());
		                        machineCuring.setMACHINE_TYPE(machineTypeValue);
		                        machineCuring.setCAVITY(BigDecimal.valueOf(cavityCell.getNumericCellValue()));
		                        machineCuring.setSTATUS_USAGE(BigDecimal.valueOf(statususage.getNumericCellValue()));
		                        machineCuring.setSTATUS(BigDecimal.valueOf(1));
		                        machineCuring.setCREATION_DATE(new Date());
		                        machineCuring.setLAST_UPDATE_DATE(new Date());

		                        machineCurings.add(machineCuring);
		                    }
		                }

		                if (!errorMessages.isEmpty()) {
		                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
		                }

		                machineCuringServiceImpl.deleteAllMachineCuring();
		                for (MachineCuring machineCuring : machineCurings) {
		                    machineCuringServiceImpl.saveMachineCuring(machineCuring);
		                }

		                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), machineCurings);
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
		
	    @RequestMapping("/exportMachineCuringExcel")
	    public ResponseEntity<InputStreamResource> exportMachineCuringExcel() throws IOException {
	        String filename = "MASTER_MACHINE_CURING.xlsx";

	        ByteArrayInputStream data = machineCuringServiceImpl.exportMachineCuringsExcel(); 
	        InputStreamResource file = new InputStreamResource(data); 

	        return ResponseEntity.ok()
	                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
	                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) 
	                .body(file); 
	    }
	    
	    @RequestMapping("/layoutMachineCuringExcel")
	    public ResponseEntity<InputStreamResource> layoutMachineCuringExcel() throws IOException {
	        String filename = "LAYOUT_MACHINE_CURING.xlsx";

	        ByteArrayInputStream data = machineCuringServiceImpl.layoutMachineCuringsExcel(); 
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

