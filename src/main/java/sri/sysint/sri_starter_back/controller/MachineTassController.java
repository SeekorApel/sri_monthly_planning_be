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
import sri.sysint.sri_starter_back.model.MachineTass;
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.repository.MachineTassTypeRepo;
import sri.sysint.sri_starter_back.service.MachineTassServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineTassController {
		
	private Response response;	

	@Autowired
	private MachineTassServiceImpl machineTassServiceImpl;
	@Autowired
	private MachineTassTypeRepo machineTassTypeRepo;
    @Autowired
    private BuildingRepo buildingRepo;
	@PersistenceContext	
	private EntityManager em;
	
	//START - GET MAPPING
		@GetMapping("/getAllMachineTass")
		public Response getAllMachineTass(final HttpServletRequest req) throws ResourceNotFoundException {
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
		        	List<MachineTass> machineTasses = new ArrayList<>();
		        	machineTasses = machineTassServiceImpl.getAllMachineTass();

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        machineTasses
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@GetMapping("/getMachineTassById/{id}")
		public Response getMachineTassById(final HttpServletRequest req, @PathVariable String id) throws ResourceNotFoundException {
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
		        	Optional<MachineTass> machineTass = Optional.of(new MachineTass());
		        	machineTass = machineTassServiceImpl.getMachineTassById(id);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        machineTass
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
		@PostMapping("/saveMachineTass")
		public Response saveMachineTass(final HttpServletRequest req, @RequestBody MachineTass machineTass) throws ResourceNotFoundException {
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
		        	MachineTass savedMachineTass = machineTassServiceImpl.saveMachineTass(machineTass);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        savedMachineTass
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/updateMachineTass")
		public Response updateMachineTass(final HttpServletRequest req, @RequestBody MachineTass machineTass) throws ResourceNotFoundException {
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
		        	MachineTass updatedMachineTass = machineTassServiceImpl.updateMachineTass(machineTass);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        updatedMachineTass
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/deleteMachineTass")
		public Response deleteMachineTass(final HttpServletRequest req, @RequestBody MachineTass machineTass) throws ResourceNotFoundException {
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
		        	MachineTass deletedMachineTass = machineTassServiceImpl.deleteMachineTass(machineTass);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        deletedMachineTass
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/restoreMachineTass")
		public Response restoreMachineTass(final HttpServletRequest req, @RequestBody MachineTass machineTass) throws ResourceNotFoundException {
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
		            MachineTass restoredMachineTass = machineTassServiceImpl.restoreMachineTass(machineTass);

		            response = new Response(
		                new Date(),
		                HttpStatus.OK.value(),
		                null,
		                HttpStatus.OK.getReasonPhrase(),
		                req.getRequestURI(),
		                restoredMachineTass
		            );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
		
		@PostMapping("/saveMachineTassExcel")
		@Transactional
		public Response saveMachineTassExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

		                List<MachineTass> machineTasses = new ArrayList<>();
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

		                        MachineTass machineTass = new MachineTass();
		                        Cell idMachineTassCell = row.getCell(1);
		                        Cell buildingNameCell = row.getCell(2);
		                        Cell floorCell = row.getCell(3);
		                        Cell machineNumberCell = row.getCell(4);
		                        Cell typeCell = row.getCell(5);
		                        Cell workCenterTextCell = row.getCell(6);

		                        if (idMachineTassCell == null || idMachineTassCell.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 2 (ID Machine Tass)");
		                            continue;
		                        }

		                        if (buildingNameCell == null || buildingNameCell.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Building Name)");
		                            continue;
		                        }

		                        if (floorCell == null || floorCell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Floor)");
		                            continue;
		                        }

		                        if (machineNumberCell == null || machineNumberCell.getCellType() != CellType.NUMERIC) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 5 (Machine Number)");
		                            continue;
		                        }

		                        if (typeCell == null || typeCell.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 6 (Type)");
		                            continue;
		                        }

		                        if (workCenterTextCell == null || workCenterTextCell.getCellType() == CellType.BLANK) {
		                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 7 (Work Center Text)");
		                            continue;
		                        }

		                        machineTass.setID_MACHINE_TASS(idMachineTassCell.getStringCellValue());

		                        String buildingName = buildingNameCell.getStringCellValue();
		                        Optional<Building> buildingOpt = buildingRepo.findByName(buildingName);

		                        Optional<MachineTassType> machineTassTypeOpt = machineTassTypeRepo.findById(typeCell.getStringCellValue());
		                        
		                        if (buildingOpt.isPresent() ) {
		                            machineTass.setBUILDING_ID(buildingOpt.get().getBUILDING_ID());
		                        } else {
		                            errorMessages.add("Data Tidak Valid, Data Building pada Baris " + (i + 1) + " Tidak Ditemukan");
		                        }
		                        
		                        if (machineTassTypeOpt.isPresent() ) {
			                        machineTass.setMACHINE_TASS_TYPE_ID(typeCell.getStringCellValue());
		                        } else {
		                            errorMessages.add("Data Tidak Valid, Data Machine Tass Type pada Baris " + (i + 1) + " Tidak Ditemukan");
		                        }
		                        
		                        machineTass.setFLOOR(BigDecimal.valueOf(floorCell.getNumericCellValue()));
		                        machineTass.setMACHINE_NUMBER(BigDecimal.valueOf(machineNumberCell.getNumericCellValue()));
		                        machineTass.setWORK_CENTER_TEXT(workCenterTextCell.getStringCellValue());
		                        machineTass.setSTATUS(BigDecimal.valueOf(1));
		                        machineTass.setCREATION_DATE(new Date());
		                        machineTass.setLAST_UPDATE_DATE(new Date());
		                        
		                        machineTasses.add(machineTass);
		                    }
		                }

		                if (!errorMessages.isEmpty()) {
		                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
		                }

		                machineTassServiceImpl.deleteAllMachineTass();
		                for (MachineTass machineTass : machineTasses) {
		                    machineTassServiceImpl.saveMachineTass(machineTass);
		                }

		                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), machineTasses);

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

		
		   @RequestMapping("/exportMachineTassExcel")
		    public ResponseEntity<InputStreamResource> exportMachineTassExcel() throws IOException {
		        String filename = "EXPORT_MASTER_MACHINE_TASS.xlsx";

		        ByteArrayInputStream data = machineTassServiceImpl.exportMachineTassExcel(); 
		        InputStreamResource file = new InputStreamResource(data); 

		        return ResponseEntity.ok()
		                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename) 
		                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) 
		                .body(file); 
		    }
		   
		   @RequestMapping("/layoutMachineTassExcel")
		    public ResponseEntity<InputStreamResource> layoutMachineTassExcel() throws IOException {
		        String filename = "LAYOUT_MASTER_MACHINE_TASS.xlsx";

		        ByteArrayInputStream data = machineTassServiceImpl.layoutMachineTassExcel(); 
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
