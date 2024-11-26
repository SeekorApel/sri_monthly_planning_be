package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.servlet.http.HttpServletRequest;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
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
import sri.sysint.sri_starter_back.model.DWorkDay;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.service.DWorkDayServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class DWorkDayController {
		
	private Response response;	

	@Autowired
	private DWorkDayServiceImpl dWorkDayServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
	@GetMapping("/getAllDWorkDay")
	public Response getAllDWorkDay(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<DWorkDay> dWorkDays = new ArrayList<>();
	    	    dWorkDays = dWorkDayServiceImpl.getAllDWorkDays();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        dWorkDays
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getDWorkDayById/{id}")
	public Response getDWorkDayById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<DWorkDay> dWorkDay = Optional.of(new DWorkDay());
	    	    dWorkDay = dWorkDayServiceImpl.getDWorkDayById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        dWorkDay
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getDWorkDayByDate/{date}")
	public Response getDWorkDayByDate(final HttpServletRequest req, @PathVariable String date) throws ResourceNotFoundException {
	    // Validate JWT token
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

	        if (user == null) {
	            throw new ResourceNotFoundException("User not found");
	        }

	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
	        Date parsedDate = dateFormat.parse(date);

	        List<DWorkDay> dworkDays = dWorkDayServiceImpl.getDWorkDayByDate(parsedDate);

	        response = new Response(
	            new Date(),
	            HttpStatus.OK.value(),
	            null,
	            HttpStatus.OK.getReasonPhrase(),
	            req.getRequestURI(),
	            dworkDays
	        );

	    } catch (Exception e) {
	        throw new ResourceNotFoundException("Error processing request: " + e.getMessage());
	    }

	    return response;
	}

	@PostMapping("/saveDWorkDay")
	public Response saveDWorkDay(final HttpServletRequest req, @RequestBody DWorkDay dWorkDay) throws ResourceNotFoundException {
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
	        	DWorkDay savedDWorkDay = dWorkDayServiceImpl.saveDWorkDay(dWorkDay);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedDWorkDay
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateDWorkDay")
	public Response updateDWorkDay(final HttpServletRequest req, @RequestBody DWorkDay dWorkDay) throws ResourceNotFoundException {
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
	        	DWorkDay updatedDWorkDay = dWorkDayServiceImpl.updateDWorkDay(dWorkDay);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedDWorkDay
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteDWorkDay")
	public Response deleteDWorkDay(final HttpServletRequest req, @RequestBody DWorkDay dWorkDay) throws ResourceNotFoundException {
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
	        	DWorkDay deletedDWorkDay = dWorkDayServiceImpl.deleteDWorkDay(dWorkDay);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedDWorkDay
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreDWorkDay")
	public Response restoreDWorkDay(final HttpServletRequest req, @RequestBody DWorkDay dWorkDay) throws ResourceNotFoundException {
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
	            DWorkDay restoredDWorkDay = dWorkDayServiceImpl.restoreDWorkDay(dWorkDay);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredDWorkDay
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	
	@PostMapping("/saveDWorkDaysExcel")
	public Response saveDWorkDaysExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
	    String header = req.getHeader("Authorization");

	    if (header == null || !header.startsWith("Bearer ")) {
	        throw new ResourceNotFoundException("JWT token not found or maybe not valid");
	    }

	    String token = header.replace("Bearer ", "");
	    Response response;

	    try {
	        String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
	            .build()
	            .verify(token)
	            .getSubject();

	        if (user != null) {
	            if (file.isEmpty()) {
	                return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "No file uploaded", req.getRequestURI(), null);
	            }

	            dWorkDayServiceImpl.deleteAllDWorkDay();

	            try (InputStream inputStream = file.getInputStream()) {
	                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	                XSSFSheet sheet = workbook.getSheetAt(0);

	                List<DWorkDay> dWorkDays = new ArrayList<>();
	                DateTimeFormatter shortMonthFormatter = DateTimeFormatter.ofPattern("EEE MMM dd, yyyy", Locale.ENGLISH);
	                DateTimeFormatter longMonthFormatter = DateTimeFormatter.ofPattern("EEE MMMM dd, yyyy", Locale.ENGLISH);

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

	                        DWorkDay dWorkDay = new DWorkDay();
	                        Cell parentCell = row.getCell(3);
	                        Cell descriptionCell = row.getCell(4);
	                        Cell dateCell = row.getCell(2);

	                        if (parentCell != null && parentCell.getCellType() == CellType.STRING &&
	                            descriptionCell != null && descriptionCell.getCellType() == CellType.STRING &&
	                            dateCell != null) {

	                            dWorkDay.setPARENT(parentCell.getStringCellValue());
	                            dWorkDay.setDESCRIPTION(descriptionCell.getStringCellValue());

	                            try {
	                                ZonedDateTime zonedDateTime;

	                                if (dateCell.getCellType() == CellType.NUMERIC) {
	                                    if (DateUtil.isCellDateFormatted(dateCell)) {
	                                        LocalDate localDate = dateCell.getDateCellValue().toInstant()
	                                                .atZone(ZoneId.systemDefault())
	                                                .toLocalDate();
	                                        zonedDateTime = localDate.atStartOfDay(ZoneId.of("UTC"));
	                                    } else {
	                                        return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null,
	                                                "Invalid date format in file: " + dateCell.toString(), req.getRequestURI(), null);
	                                    }
	                                } else if (dateCell.getCellType() == CellType.STRING) {
	                                    String dateStr = dateCell.getStringCellValue().trim();

	                                    // Try parsing with short month format first, then with long month format
	                                    LocalDate parsedDate;
	                                    try {
	                                        parsedDate = LocalDate.parse(dateStr, shortMonthFormatter);
	                                    } catch (DateTimeParseException e) {
	                                        parsedDate = LocalDate.parse(dateStr, longMonthFormatter);
	                                    }

	                                    zonedDateTime = parsedDate.atStartOfDay(ZoneId.of("UTC"));
	                                } else {
	                                    continue;
	                                }

	                                dWorkDay.setDATE_WD(Date.from(zonedDateTime.toInstant()));
	                                dWorkDay.setDETAIL_WD_ID(dWorkDayServiceImpl.getNewId());
	                                dWorkDay.setSTATUS(BigDecimal.valueOf(1));
	                                dWorkDay.setCREATION_DATE(new Date());
	                                dWorkDay.setLAST_UPDATE_DATE(new Date());

	                                dWorkDayServiceImpl.saveDWorkDay(dWorkDay);
	                                dWorkDays.add(dWorkDay);

	                            } catch (Exception e) {
	                                return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null,
	                                        "Error parsing date: " + dateCell.toString(), req.getRequestURI(), null);
	                            }
	                        } else {
	                            System.out.println("Row " + i + " has invalid data: " +
	                                    "Parent cell: " + (parentCell != null ? parentCell.getStringCellValue() : "null") + ", " +
	                                    "Description cell: " + (descriptionCell != null ? descriptionCell.getStringCellValue() : "null") + ", " +
	                                    "Date cell: " + (dateCell != null ? dateCell.toString() : "null"));
	                        }
	                    }
	                }

	                response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), dWorkDays);

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




	
	@RequestMapping("/exportDWorkDaysExcel")
	public ResponseEntity<InputStreamResource> exportDWorkDaysExcel() throws IOException {
	    String filename = "MASTER_D_WORK_DAY.xlsx";
	    
	    ByteArrayInputStream data = dWorkDayServiceImpl.exportDWorkDaysExcel();
	    InputStreamResource file = new InputStreamResource(data);
	    
	    return ResponseEntity.ok()
	        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
	        .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
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
