package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

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
import sri.sysint.sri_starter_back.model.CTAssy;
import sri.sysint.sri_starter_back.model.CTCuring;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.CTCuringServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class CTCuringController {
	
	private Response response;	

	@Autowired
	private CTCuringServiceImpl ctCuringServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
	@GetMapping("/getAllCTCuring")
	public Response getAllCTCuring(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<CTCuring> ctCurings = new ArrayList<>();
	        	ctCurings = ctCuringServiceImpl.getAllCTCuring();
		
			    response = new Response(
			        new Date(),
			        HttpStatus.OK.value(),
			        null,
			        HttpStatus.OK.getReasonPhrase(),
			        req.getRequestURI(),
			        ctCurings
			    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getCTCuringById/{id}")
	public Response getCTCuringById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<CTCuring> ctCuring = Optional.of(new CTCuring());
	        	ctCuring = ctCuringServiceImpl.getCTCuringById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        ctCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveCTCuring")
	public Response saveCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring savedCTCuring = ctCuringServiceImpl.saveCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateCTCuring")
	public Response updateCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring updatedCTCuring = ctCuringServiceImpl.updateCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteCTCuring")
	public Response deleteCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring deletedCTCuring = ctCuringServiceImpl.deleteCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreCTCuring")
	public Response activateCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring activatedCTCuring = ctCuringServiceImpl.activateCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        activatedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveCTCuringExcel")
	public Response saveCTCuringExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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
	    	    
	    	    ctCuringServiceImpl.deleteAllCTCuring();

	    	    try (InputStream inputStream = file.getInputStream()) {
	    	        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	    	        XSSFSheet sheet = workbook.getSheetAt(0);
	    	        
	    	        List<CTCuring> ctCurings = new ArrayList<>();

//	    	        boolean isTrue = true;
//	    	        Row headerRow = sheet.getRow(0);
//	    	        Row headerRow1 = sheet.getRow(1);
//	    	        Row headerRow2 = sheet.getRow(2);
//	    	        
//	    	        if(headerRow != null) {
//	    	        	  for (int i = 1; i < 37; i++) {
//	    	        		  Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
//	    	        		  
//	    	        		   if (cell != null) {
//	    	        			   if (!ctAssyServiceImpl.getCellValue(headerRow1.getCell(1),CTAssyServiceImpl.Type.STRING).replace(".0", "").equals("NO")) {
//	    	                            isTrue = false; 
//	    	                        }
//	    	        		   }
//	    	        	  }
//	    	        }
	    	        
	    	        for (int i = 3; i <= sheet.getLastRowNum(); i++) {
	    	            Row row = sheet.getRow(i);
	    	            if (row != null) {
	    	                CTCuring ctCuring = new CTCuring();

	    	                Cell nameCell = row.getCell(1);

	    	                if (nameCell != null) {
	    	                    ctCuring.setCT_CURING_ID(ctCuringServiceImpl.getNewId());  
	    	                    ctCuring.setWIP(getStringFromCell(row.getCell(0)));  // Column A
	    	                    ctCuring.setGROUP_COUNTER(getStringFromCell(row.getCell(1)));  // Column C
	    	                    ctCuring.setVAR_GROUP_COUNTER(getStringFromCell(row.getCell(2)));  // Column D
	    	                    ctCuring.setSEQUENCE(getBigDecimalFromCell(row.getCell(3)));  // Column E
	    	                    ctCuring.setWCT(getStringFromCell(row.getCell(4)));  // Column F
	    	                    ctCuring.setOPERATION_SHORT_TEXT(getStringFromCell(row.getCell(5)));  // Column G
	    	                    ctCuring.setOPERATION_UNIT(getStringFromCell(row.getCell(6)));  // Column H
	    	                    ctCuring.setBASE_QUANTITY(getBigDecimalFromCell(row.getCell(7)));  // Column I
	    	                    ctCuring.setSTANDART_VALUE_UNIT(getStringFromCell(row.getCell(8)));  // Column J
	    	                    ctCuring.setCT_SEC1(getBigDecimalFromCell(row.getCell(9))); // Column K
	    	                    ctCuring.setCT_HR1000(getBigDecimalFromCell(row.getCell(10)));  // Column L
	    	                    ctCuring.setWH_NORMAL_SHIFT_0(getBigDecimalFromCell(row.getCell(11))); // Column M
	    	                    ctCuring.setWH_NORMAL_SHIFT_1(getBigDecimalFromCell(row.getCell(12)));  // Column N
	    	                    ctCuring.setWH_NORMAL_SHIFT_2(getBigDecimalFromCell(row.getCell(13))); // Column O
	    	                    ctCuring.setWH_SHIFT_FRIDAY(getBigDecimalFromCell(row.getCell(14)));  // Renamed JUMAT to FRIDAY (Column P)
	    	                    ctCuring.setWH_TOTAL_NORMAL_SHIFT(getBigDecimalFromCell(row.getCell(15)));  // Column Q
	    	                    ctCuring.setWH_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(row.getCell(16)));  // Renamed JUMAT to FRIDAY (Column R)
	    	                    ctCuring.setALLOW_NORMAL_SHIFT_0(getBigDecimalFromCell(row.getCell(17)));  // Column S
	    	                    ctCuring.setALLOW_NORMAL_SHIFT_1(getBigDecimalFromCell(row.getCell(18)));  // Column T
	    	                    ctCuring.setALLOW_NORMAL_SHIFT_2(getBigDecimalFromCell(row.getCell(19)));  // Column U
	    	                    ctCuring.setALLOW_TOTAL(getBigDecimalFromCell(row.getCell(20)));  // Column V
	    	                    ctCuring.setOP_TIME_NORMAL_SHIFT_0(getBigDecimalFromCell(row.getCell(21)));  // Column W
	    	                    ctCuring.setOP_TIME_NORMAL_SHIFT_1(getBigDecimalFromCell(row.getCell(22)));  // Column X
	    	                    ctCuring.setOP_TIME_NORMAL_SHIFT_2(getBigDecimalFromCell(row.getCell(23)));  // Column Y
	    	                    ctCuring.setOP_TIME_SHIFT_FRIDAY(getBigDecimalFromCell(row.getCell(24))); // Renamed JUMAT to FRIDAY (Column Z)
	    	                    ctCuring.setOP_TIME_NORMAL_SHIFT(getBigDecimalFromCell(row.getCell(25)));  // Column AA
	    	                    ctCuring.setOP_TIME_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(row.getCell(26))); // Renamed JUMAT to FRIDAY (Column AB)
	    	                    ctCuring.setKAPS_NORMAL_SHIFT_0(getBigDecimalFromCell(row.getCell(27)));  // Column AC
	    	                    ctCuring.setKAPS_NORMAL_SHIFT_1(getBigDecimalFromCell(row.getCell(28))); // Column AD
	    	                    ctCuring.setKAPS_NORMAL_SHIFT_2(getBigDecimalFromCell(row.getCell(29))); // Column AE
	    	                    ctCuring.setKAPS_SHIFT_FRIDAY(getBigDecimalFromCell(row.getCell(30))); // Renamed JUMAT to FRIDAY (Column AF)
	    	                    ctCuring.setKAPS_TOTAL_NORMAL_SHIFT(getBigDecimalFromCell(row.getCell(31))); // Column AG
	    	                    ctCuring.setKAPS_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(row.getCell(32))); // Renamed JUMAT to FRIDAY (Column AH)
	    	                    ctCuring.setWAKTU_TOTAL_CT_NORMAL(getBigDecimalFromCell(row.getCell(33))); // Column AI
	    	                    ctCuring.setWAKTU_TOTAL_CT_FRIDAY(getBigDecimalFromCell(row.getCell(34)));  // Renamed JUMAT to FRIDAY (Column AJ)
	    	                    ctCuring.setSTATUS(BigDecimal.valueOf(1));  
	    	                    ctCuring.setCREATION_DATE(new Date());  
	    	                    ctCuring.setLAST_UPDATE_DATE(new Date());  
	    	                }


	    	                ctCuringServiceImpl.saveCTCuring(ctCuring);
	    	                ctCurings.add(ctCuring);
	    	            }
	    	        }

	    	        response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), ctCurings);

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
	                return null; 
	            }
	        case NUMERIC:
	            return BigDecimal.valueOf(cell.getNumericCellValue());
	        default:
	            return null; 
	    }
	}

	private String getStringFromCell(Cell cell) {
	    if (cell == null) {
	        return null;
	    }
	    return cell.getStringCellValue();
	}
	
    @GetMapping("/exportCTCuringExcel")
    public ResponseEntity<InputStreamResource> exportCTCuringExcel() throws IOException {
        String filename = "MASTER_CT_CURING_DATA.xlsx";

        // Generate the Excel data using the service method
        ByteArrayInputStream data = ctCuringServiceImpl.exportCTCuringsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    
}
