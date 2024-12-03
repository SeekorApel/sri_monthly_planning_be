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
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.MachineCuringTypeServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineCuringTypeController {
		
	private Response response;	

	@Autowired
	private MachineCuringTypeServiceImpl machineCuringTypeServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/getAllMachineCuringType")
	public Response getAllPlant(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<MachineCuringType> machineCuringTypes = new ArrayList<>();
	        	machineCuringTypes = machineCuringTypeServiceImpl.getAllMachineCuringType();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        machineCuringTypes
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getMachineCuringTypeById/{id}")
	public Response getPlantById(final HttpServletRequest req, @PathVariable String id) throws ResourceNotFoundException {
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
	        	Optional<MachineCuringType> machineCuringType = Optional.of(new MachineCuringType());
	        	machineCuringType = machineCuringTypeServiceImpl.getMachineCuringTypeById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        machineCuringType
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
	@PostMapping("/saveMachineCuringType")
	public Response savePlant(final HttpServletRequest req, @RequestBody MachineCuringType machineCuringType) throws ResourceNotFoundException {
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
	        	MachineCuringType savedMachineCuringType = machineCuringTypeServiceImpl.saveMachineCuringType(machineCuringType);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedMachineCuringType
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateMachineCuringType")
	public Response updatePlant(final HttpServletRequest req, @RequestBody MachineCuringType machineCuringType) throws ResourceNotFoundException {
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
	        	MachineCuringType updatedMachineCuringType = machineCuringTypeServiceImpl.updateMachineCuringType(machineCuringType);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedMachineCuringType
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteMachineCuringType")
	public Response deletetePlant(final HttpServletRequest req, @RequestBody MachineCuringType machineCuringType) throws ResourceNotFoundException {
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
	        	MachineCuringType deletedMachineCuringType = machineCuringTypeServiceImpl.deleteMachineCuringType(machineCuringType);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedMachineCuringType
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreMachineCuringType")
	public Response restoreMachineCuringType(final HttpServletRequest req, @RequestBody MachineCuringType machineCuringType) throws ResourceNotFoundException {
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
	            MachineCuringType restoredMachineCuringType = machineCuringTypeServiceImpl.restoreMachineCuringType(machineCuringType);

	            Response response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredMachineCuringType
	            );
	            return response;
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }
	}

	
	@PostMapping("/saveMachineCuringTypeExcel")
	public Response saveMachineCuringTypesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	            machineCuringTypeServiceImpl.deleteAllMachineCuringType();

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);

	                List<MachineCuringType> machineCuringTypes = new ArrayList<>();

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

	                        MachineCuringType machineCuringType = new MachineCuringType();

	                        Cell machineCuringTypeIdCell = row.getCell(1);
	                        Cell settingIdCell = row.getCell(2);
	                        Cell descriptionCell = row.getCell(3);
	                        Cell cavityCell = row.getCell(4);

	                        if (settingIdCell != null && settingIdCell.getCellType() == CellType.NUMERIC
	                                && descriptionCell != null 
	                                && cavityCell != null && cavityCell.getCellType() == CellType.NUMERIC) {

	                            machineCuringType.setMACHINECURINGTYPE_ID(machineCuringTypeIdCell.getStringCellValue());
	                            machineCuringType.setSETTING_ID(BigDecimal.valueOf(settingIdCell.getNumericCellValue()));
	                            machineCuringType.setDESCRIPTION(descriptionCell.getStringCellValue());
	                            machineCuringType.setCAVITY(BigDecimal.valueOf(cavityCell.getNumericCellValue()));
	                            machineCuringType.setSTATUS(BigDecimal.valueOf(1));
	                            machineCuringType.setCREATION_DATE(new Date());
	                            machineCuringType.setLAST_UPDATE_DATE(new Date());

	                            machineCuringTypeServiceImpl.saveMachineCuringType(machineCuringType);
	                            machineCuringTypes.add(machineCuringType);
	                        } else {
	                            continue;
	                        }
	                    }
	                }

	                response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), machineCuringTypes);

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

    @GetMapping("/exportMachineCuringTypeexcel")
    public ResponseEntity<InputStreamResource> exportMachineCuringTypesExcel() throws IOException {
        String filename = "MASTER_MACHINE_CURING_TYPE.xlsx";

        ByteArrayInputStream data = machineCuringTypeServiceImpl.exportMachineCuringTypesExcel();
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

