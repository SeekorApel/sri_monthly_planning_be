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
import sri.sysint.sri_starter_back.model.CtKapa;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.CtKapaServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class CtKapaController {
		
	private Response response;	

	@Autowired
	private CtKapaServiceImpl ctKapaServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/getAllCtKapa")
	public Response getAllCtKapa(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<CtKapa> ctKapas = new ArrayList<>();
	    	    ctKapas = ctKapaServiceImpl.getAllCtKapa();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        ctKapas
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getCtKapaById/{id}")
	public Response getCtKapaById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<CtKapa> ctKapa = Optional.of(new CtKapa());
	    	    ctKapa = ctKapaServiceImpl.getCtKapaById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        ctKapa
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
	@PostMapping("/saveCtKapa")
	public Response saveCtKapa(final HttpServletRequest req, @RequestBody CtKapa ctKapa) throws ResourceNotFoundException {
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
	        	CtKapa savedCtKapa = ctKapaServiceImpl.saveCtKapa(ctKapa);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedCtKapa
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateCtKapa")
	public Response updateCtKapa(final HttpServletRequest req, @RequestBody CtKapa ctKapa) throws ResourceNotFoundException {
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
	        	CtKapa updatedCtKapa = ctKapaServiceImpl.updateCtKapa(ctKapa);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedCtKapa
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteCtKapa")
	public Response deleteteCtKapa(final HttpServletRequest req, @RequestBody CtKapa ctKapa) throws ResourceNotFoundException {
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
	        	CtKapa deletedCtKapa = ctKapaServiceImpl.deleteCtKapa(ctKapa);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedCtKapa
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreCtKapa")
	public Response restoreCtKapa(final HttpServletRequest req, @RequestBody CtKapa ctKapa) throws ResourceNotFoundException {
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
	            CtKapa restoredCtKapa = ctKapaServiceImpl.restoreCtKapa(ctKapa);

	            Response response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredCtKapa
	            );
	            return response;
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }
	}

	
	@PostMapping("/saveCtKapasExcel")
	public Response saveCtKapasExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	            ctKapaServiceImpl.deleteAllCtKapa();

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);

	                List<CtKapa> ctKapas = new ArrayList<>();

	                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);
	                    if (row != null) {
	                        CtKapa ctKapa = new CtKapa();

	                        Cell itemCuringCell = row.getCell(2);
	                        Cell typeCuringCell = row.getCell(3);
	                        Cell descriptionCell = row.getCell(4);
	                        Cell cycleTimeCell = row.getCell(5);
	                        Cell shiftCell = row.getCell(6);
	                        Cell kapaPerShiftCell = row.getCell(7);
	                        Cell lastUpdateDataCell = row.getCell(8);
	                        Cell machineCell = row.getCell(9);

	                        if (itemCuringCell != null && itemCuringCell.getCellType() == CellType.STRING &&
	                            typeCuringCell != null && typeCuringCell.getCellType() == CellType.STRING &&
	                            descriptionCell != null && descriptionCell.getCellType() == CellType.STRING &&
	                            cycleTimeCell != null && cycleTimeCell.getCellType() == CellType.NUMERIC &&
	                            shiftCell != null && shiftCell.getCellType() == CellType.NUMERIC &&
	                            kapaPerShiftCell != null && kapaPerShiftCell.getCellType() == CellType.NUMERIC &&
	                            lastUpdateDataCell != null && lastUpdateDataCell.getCellType() == CellType.NUMERIC &&
	                            machineCell != null && machineCell.getCellType() == CellType.STRING) {

	                            ctKapa.setITEM_CURING(itemCuringCell.getStringCellValue());
	                            ctKapa.setTYPE_CURING(typeCuringCell.getStringCellValue());
	                            ctKapa.setDESCRIPTION(descriptionCell.getStringCellValue());
	                            ctKapa.setCYCLE_TIME(BigDecimal.valueOf(cycleTimeCell.getNumericCellValue()));
	                            ctKapa.setSHIFT(BigDecimal.valueOf(shiftCell.getNumericCellValue()));
	                            ctKapa.setKAPA_PERSHIFT(BigDecimal.valueOf(kapaPerShiftCell.getNumericCellValue()));
	                            ctKapa.setLAST_UPDATE_DATA(BigDecimal.valueOf(lastUpdateDataCell.getNumericCellValue()));
	                            ctKapa.setMACHINE(machineCell.getStringCellValue());
	                            
	                            ctKapa.setSTATUS(BigDecimal.valueOf(1));
	                            ctKapa.setCREATION_DATE(new Date());
	                            ctKapa.setLAST_UPDATE_DATE(new Date());

	                            ctKapaServiceImpl.saveCtKapa(ctKapa);
	                            ctKapas.add(ctKapa);
	                        } else {
	                        	continue;
	                        }
	                    }
	                }

	                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), ctKapas);

	            } catch (IOException e) {
	                return new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file", req.getRequestURI(), null);
	            }
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }
	}
	
	@GetMapping("/exportCtKapaExcel")
    public ResponseEntity<InputStreamResource> exportCtKapaExcel() throws IOException {
        String filename = "MASTER_CT_KAPA.xlsx";

        ByteArrayInputStream data = ctKapaServiceImpl.exportCtKapasExcel();
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


