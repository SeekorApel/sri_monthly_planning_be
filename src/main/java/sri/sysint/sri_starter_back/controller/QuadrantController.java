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
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.Quadrant;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.service.QuadrantServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class QuadrantController {
		
	private Response response;	

	@Autowired
	private QuadrantServiceImpl quadrantServiceImpl;
	@Autowired
    private BuildingRepo buildingRepo;
	
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/getAllQuadrant")
	public Response getAllQuadrant(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<Quadrant> quadrants = new ArrayList<>();
	        	quadrants = quadrantServiceImpl.getAllQuadrant();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        quadrants
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getQuadrantById/{id}")
	public Response getQuadrantById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<Quadrant> quadrant = Optional.of(new Quadrant());
	        	quadrant = quadrantServiceImpl.getQuadrantById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        quadrant
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
	@PostMapping("/saveQuadrant")
	public Response saveQuadrant(final HttpServletRequest req, @RequestBody Quadrant quadrant) throws ResourceNotFoundException {
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
	        	Quadrant savedQuadrant = quadrantServiceImpl.saveQuadrant(quadrant);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedQuadrant
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateQuadrant")
	public Response updateQuadrant(final HttpServletRequest req, @RequestBody Quadrant quadrant) throws ResourceNotFoundException {
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
	        	Quadrant updatedQuadrant = quadrantServiceImpl.updateQuadrant(quadrant);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedQuadrant
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteQuadrant")
	public Response deleteQuadrant(final HttpServletRequest req, @RequestBody Quadrant quadrant) throws ResourceNotFoundException {
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
	        	Quadrant deletedQuadrant = quadrantServiceImpl.deleteQuadrant(quadrant);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedQuadrant
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreQuadrant")
	public Response restoreQuadrant(final HttpServletRequest req, @RequestBody Quadrant quadrant) throws ResourceNotFoundException {
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
	            Quadrant restoredQuadrant = quadrantServiceImpl.restoreQuadrant(quadrant);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredQuadrant
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	
	@PostMapping("/saveQuandrantExcel")
	@Transactional
	public Response saveQuadrantsExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	                List<Quadrant> quadrants = new ArrayList<>();
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

	                       CT_Curing ctCuring = new CT_Curing();
boolean hasError = false;

// Loop through columns to validate cells
for (int col = 0; col <= 34; col++) {
    Cell cell = row.getCell(col);
    if (cell == null || cell.getCellType() == CellType.BLANK) {
        errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom " + (col + 1));
        hasError = true;
    }
}

// Skip processing the row if there are errors
if (hasError) {
    continue;
}

// Proceed with processing the row
Cell wipCell = row.getCell(0);
Cell groupCounterCell = row.getCell(1);

String wip = wipCell.getStringCellValue();
Optional<CT_WIP> wipOpt = wipRepo.findByWIP(wip);

if (wipOpt.isPresent()) {
    CT_WIP ctWip = wipOpt.get();
    ctCuring.setCT_CURING_ID(ctCuringServiceImpl.getNewId());
    ctCuring.setWIP(wip);
    ctCuring.setGROUP_COUNTER(groupCounterCell.getStringCellValue());

    for (int col = 2; col <= 34; col++) {
        Cell cell = row.getCell(col);
        switch (col) {
            case 2:
                ctCuring.setVAR_GROUP_COUNTER(getStringFromCell(cell));
                break;
            case 3:
                ctCuring.setSEQUENCE(getBigDecimalFromCell(cell));
                break;
            case 4:
                ctCuring.setWCT(getStringFromCell(cell));
                break;
            case 5:
                ctCuring.setOPERATION_SHORT_TEXT(getStringFromCell(cell));
                break;
            case 6:
                ctCuring.setOPERATION_UNIT(getStringFromCell(cell));
                break;
            case 7:
                ctCuring.setBASE_QUANTITY(getBigDecimalFromCell(cell));
                break;
            case 8:
                ctCuring.setSTANDART_VALUE_UNIT(getStringFromCell(cell));
                break;
            case 9:
                ctCuring.setCT_SEC1(getBigDecimalFromCell(cell));
                break;
            case 10:
                ctCuring.setCT_HR1000(getBigDecimalFromCell(cell));
                break;
            case 11:
                ctCuring.setWH_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
                break;
            case 12:
                ctCuring.setWH_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
                break;
            case 13:
                ctCuring.setWH_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
                break;
            case 14:
                ctCuring.setWH_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
                break;
            case 15:
                ctCuring.setWH_TOTAL_NORMAL_SHIFT(getBigDecimalFromCell(cell));
                break;
            case 16:
                ctCuring.setWH_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
                break;
            case 17:
                ctCuring.setALLOW_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
                break;
            case 18:
                ctCuring.setALLOW_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
                break;
            case 19:
                ctCuring.setALLOW_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
                break;
            case 20:
                ctCuring.setALLOW_TOTAL(getBigDecimalFromCell(cell));
                break;
            case 21:
                ctCuring.setOP_TIME_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
                break;
            case 22:
                ctCuring.setOP_TIME_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
                break;
            case 23:
                ctCuring.setOP_TIME_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
                break;
            case 24:
                ctCuring.setOP_TIME_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
                break;
            case 25:
                ctCuring.setOP_TIME_NORMAL_SHIFT(getBigDecimalFromCell(cell));
                break;
            case 26:
                ctCuring.setOP_TIME_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
                break;
            case 27:
                ctCuring.setKAPS_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
                break;
            case 28:
                ctCuring.setKAPS_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
                break;
            case 29:
                ctCuring.setKAPS_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
                break;
            case 30:
                ctCuring.setKAPS_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
                break;
            case 31:
                ctCuring.setKAPS_TOTAL_NORMAL_SHIFT(getBigDecimalFromCell(cell));
                break;
            case 32:
                ctCuring.setKAPS_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
                break;
            case 33:
                ctCuring.setWAKTU_TOTAL_CT_NORMAL(getBigDecimalFromCell(cell));
                break;
            case 34:
                ctCuring.setWAKTU_TOTAL_CT_FRIDAY(getBigDecimalFromCell(cell));
                break;
        }
    }

    ctCuring.setSTATUS(BigDecimal.valueOf(1));
    ctCuring.setCREATION_DATE(new Date());
    ctCuring.setLAST_UPDATE_DATE(new Date());

    ctCurings.add(ctCuring);
} else {
    errorMessages.add("Data Tidak Valid, Data WIP pada Baris " + (i + 1) + " Tidak Ditemukan");
}

	                    }
	                }

	                if (!errorMessages.isEmpty()) {
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
	                }

	                quadrantServiceImpl.deleteAllQuadrant();
	                for (Quadrant quadrant : quadrants) {
	                    quadrantServiceImpl.saveQuadrant(quadrant);
	                }

	                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), quadrants);

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




    @GetMapping("/exportQuadrantsExcel")
    public ResponseEntity<InputStreamResource> exportQuadrantsExcel() throws IOException {
        String filename = "EXPORT_MASTER_QUADRANT.xlsx";
        ByteArrayInputStream data = quadrantServiceImpl.exportQuadrantsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    @GetMapping("/layoutQuadrantsExcel")
    public ResponseEntity<InputStreamResource> layoutQuadrantsExcel() throws IOException {
        String filename = "LAYOUT_MASTER_QUADRANT.xlsx";
        ByteArrayInputStream data = quadrantServiceImpl.layoutQuadrantsExcel();
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
