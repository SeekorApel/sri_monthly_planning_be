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
import sri.sysint.sri_starter_back.model.Quadrant;
import sri.sysint.sri_starter_back.model.QuadrantDistance;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.QuadrantRepo;
import sri.sysint.sri_starter_back.service.QuadrantDistanceServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class QuadrantDistanceController {
		
	private Response response;	

	@Autowired
	private QuadrantDistanceServiceImpl quadrantDistanceServiceImpl;
	
    @Autowired
    private QuadrantRepo quadrantRepo;
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/getAllQuadrantDistance")
	public Response getAllQuadrantDistance(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<QuadrantDistance> quadrantDistances = new ArrayList<>();
	    	    quadrantDistances = quadrantDistanceServiceImpl.getAllQuadrantDistance();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        quadrantDistances
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getQuadrantDistanceById/{id}")
	public Response getQuadrantDistanceById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<QuadrantDistance> quadrantDistance = Optional.of(new QuadrantDistance());
	    	    quadrantDistance = quadrantDistanceServiceImpl.getQuadrantDistanceById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        quadrantDistance
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
	@PostMapping("/saveQuadrantDistance")
	public Response saveQuadrantDistance(final HttpServletRequest req, @RequestBody QuadrantDistance quadrantDistance) throws ResourceNotFoundException {
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
	        	QuadrantDistance savedQuadrantDistance = quadrantDistanceServiceImpl.saveQuadrantDistance(quadrantDistance);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedQuadrantDistance
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateQuadrantDistance")
	public Response updateQuadrantDistance(final HttpServletRequest req, @RequestBody QuadrantDistance quadrantDistance) throws ResourceNotFoundException {
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
	            QuadrantDistance updatedQuadrantDistance = quadrantDistanceServiceImpl.updateQuadrantDistance(quadrantDistance);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                updatedQuadrantDistance
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	
	@PostMapping("/deleteQuadrantDistance")
	public Response deleteQuadrantDistance(final HttpServletRequest req, @RequestBody QuadrantDistance quadrantDistance) throws ResourceNotFoundException {
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
	            QuadrantDistance deletedQuadrantDistance = quadrantDistanceServiceImpl.deleteQuadrantDistance(quadrantDistance);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                deletedQuadrantDistance
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	@PostMapping("/restoreQuadrantDistance")
	public Response restoreQuadrantDistance(final HttpServletRequest req, @RequestBody QuadrantDistance quadrantDistance) throws ResourceNotFoundException {
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
	            QuadrantDistance restoredQuadrantDistance = quadrantDistanceServiceImpl.restoreQuadrantDistance(quadrantDistance);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredQuadrantDistance
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	
	@PostMapping("/saveQuadrantDistancesExcel")
	public Response saveQuadrantDistancesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	            quadrantDistanceServiceImpl.deleteAllQuadrantDistance();

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);

	                List<QuadrantDistance> quadrantDistances = new ArrayList<>();

	                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);

	                    if (row != null) {
	                        boolean isEmptyRow = true;

	                        // Check if row is empty
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

	                        QuadrantDistance quadrantDistance = new QuadrantDistance();

	                        Cell quadrantName1Cell = row.getCell(2);
	                        Cell quadrantName2Cell = row.getCell(3);
	                        Cell distanceCell = row.getCell(4);

	                        if (quadrantName1Cell != null && quadrantName1Cell.getCellType() == CellType.STRING
	                                && quadrantName2Cell != null && quadrantName2Cell.getCellType() == CellType.STRING
	                                && distanceCell != null && distanceCell.getCellType() == CellType.NUMERIC) {

	                            String quadrantName1 = quadrantName1Cell.getStringCellValue();
	                            Optional<Quadrant> quadrant1Opt = quadrantRepo.findByName(quadrantName1);

	                            if (quadrant1Opt.isPresent()) {
	                                quadrantDistance.setQUADRANT_ID_1(quadrant1Opt.get().getQUADRANT_ID());
	                            } else {
	                                quadrantDistance.setQUADRANT_ID_1(BigDecimal.ZERO);
	                            }

	                            String quadrantName2 = quadrantName2Cell.getStringCellValue();
	                            Optional<Quadrant> quadrant2Opt = quadrantRepo.findByName(quadrantName2);
	                            if (quadrant2Opt.isPresent()) {
	                                quadrantDistance.setQUADRANT_ID_2(quadrant2Opt.get().getQUADRANT_ID());
	                            } else {
	                                quadrantDistance.setQUADRANT_ID_2(BigDecimal.ZERO);
	                            }

	                            // Set the distance
	                            quadrantDistance.setDISTANCE(BigDecimal.valueOf(distanceCell.getNumericCellValue()));
	                            quadrantDistance.setID_Q_DISTANCE(quadrantDistanceServiceImpl.getNewId());
	                            quadrantDistance.setSTATUS(BigDecimal.valueOf(1));
	                            quadrantDistance.setCREATION_DATE(new Date());
	                            quadrantDistance.setLAST_UPDATE_DATE(new Date());

	                            quadrantDistanceServiceImpl.saveQuadrantDistance(quadrantDistance);
	                            quadrantDistances.add(quadrantDistance);
	                        } else {
	                            continue;
	                        }
	                    }
	                }

	                response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), quadrantDistances);

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



    @GetMapping("/exportQuadrantDistancesExcel") 
    public ResponseEntity<InputStreamResource> exportQuadrantDistancesExcel() throws IOException {
        String filename = "MASTER_QUADRANT_DISTANCE.xlsx"; 

        ByteArrayInputStream data = quadrantDistanceServiceImpl.exportQuadrantDistancesExcel();
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
