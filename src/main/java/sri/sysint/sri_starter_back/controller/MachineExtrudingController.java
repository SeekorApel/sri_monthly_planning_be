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
import sri.sysint.sri_starter_back.model.MachineExtruding;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.MachineExtrudingServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineExtrudingController {
		
	private Response response;	

	@Autowired
	private MachineExtrudingServiceImpl machineExtrudingServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/getAllMachineExtruding")
	public Response getAllMachineExtruding(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<MachineExtruding> machineExtrudings = new ArrayList<>();
	    	    machineExtrudings = machineExtrudingServiceImpl.getAllMachineExtruding();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        machineExtrudings
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getMachineExtrudingById/{id}")
	public Response getMachineExtrudingById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<MachineExtruding> machineExtruding = Optional.of(new MachineExtruding());
	    	    machineExtruding = machineExtrudingServiceImpl.getMachineExtrudingById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        machineExtruding
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
	@PostMapping("/saveMachineExtruding")
	public Response saveMachineExtruding(final HttpServletRequest req, @RequestBody MachineExtruding machineExtruding) throws ResourceNotFoundException {
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
	        	MachineExtruding savedMachineExtruding = machineExtrudingServiceImpl.saveMachineExtruding(machineExtruding);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedMachineExtruding
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateMachineExtruding")
	public Response updateMachineExtruding(final HttpServletRequest req, @RequestBody MachineExtruding machineExtruding) throws ResourceNotFoundException {
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
	        	MachineExtruding updatedMachineExtruding = machineExtrudingServiceImpl.updateMachineExtruding(machineExtruding);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedMachineExtruding
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteMachineExtruding")
	public Response deleteMachineExtruding(final HttpServletRequest req, @RequestBody MachineExtruding machineExtruding) throws ResourceNotFoundException {
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
	        	MachineExtruding deletedMachineExtruding = machineExtrudingServiceImpl.deleteMachineExtruding(machineExtruding);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedMachineExtruding
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreMachineExtruding")
	public Response restoreMachineExtruding(final HttpServletRequest req, @RequestBody MachineExtruding machineExtruding) throws ResourceNotFoundException {
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
	            MachineExtruding restoredMachineExtruding = machineExtrudingServiceImpl.restoreMachineExtruding(machineExtruding);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredMachineExtruding
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	
	@PostMapping("/saveMachineExtrudingsExcel")
	public Response saveMachineExtrudingsExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	            machineExtrudingServiceImpl.deleteAllMachineExtruding();

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);

	                List<MachineExtruding> machineExtrudings = new ArrayList<>();

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

	                        MachineExtruding machineExtruding = new MachineExtruding();
	                        Cell buildingIdCell = row.getCell(2);
	                        Cell typeCell = row.getCell(3);

	                        if (buildingIdCell != null && buildingIdCell.getCellType() == CellType.NUMERIC
	                                && typeCell != null ) {
	                            machineExtruding.setID_MACHINE_EXT(machineExtrudingServiceImpl.getNewId());
	                            machineExtruding.setBUILDING_ID(BigDecimal.valueOf(buildingIdCell.getNumericCellValue()));
	                            machineExtruding.setTYPE(typeCell.getStringCellValue());
	                            machineExtruding.setSTATUS(BigDecimal.valueOf(1));
	                            machineExtruding.setCREATION_DATE(new Date());
	                            machineExtruding.setLAST_UPDATE_DATE(new Date());

	                            machineExtrudingServiceImpl.saveMachineExtruding(machineExtruding);
	                            machineExtrudings.add(machineExtruding);
	                        } else {
	                            continue;
	                        }
	                    }
	                }

	                response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), machineExtrudings);

	            } catch (IOException e) {
	                response = new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file", req.getRequestURI(), null);
	            }
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

    @RequestMapping("/exportMachineExtrudingExcel")
    public ResponseEntity<InputStreamResource> exportMachineExtrudingExcel() throws IOException {
        String filename = "MASTER_MACHINE_EXTRUDING.xlsx"; 

        ByteArrayInputStream data = machineExtrudingServiceImpl.exportMachineExtrudingsExcel();
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