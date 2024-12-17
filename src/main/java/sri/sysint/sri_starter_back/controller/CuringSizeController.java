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
import sri.sysint.sri_starter_back.model.CuringSize;
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.repository.MachineCuringTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;
import sri.sysint.sri_starter_back.service.CuringSizeServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class CuringSizeController {

	private Response response;	
	@Autowired
    private MachineCuringTypeRepo machineCuringTypeRepo;
	@Autowired
	private SizeRepo sizeRepo;
	@Autowired
	private CuringSizeServiceImpl curingSizeServiceImpl;
	
	@PersistenceContext	
	private EntityManager em;
	
	@GetMapping("/getAllCuringSize")
	public Response getAllCuringSize(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<CuringSize> curingSizes = new ArrayList<>();
	    		curingSizes = curingSizeServiceImpl.getAllCuringSize();
		
			    response = new Response(
			        new Date(),
			        HttpStatus.OK.value(),
			        null,
			        HttpStatus.OK.getReasonPhrase(),
			        req.getRequestURI(),
			        curingSizes
			    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	
	@GetMapping("/getCuringSizeById/{id}")
	public Response getCuringSizeById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<CuringSize> curingSize = Optional.of(new CuringSize());
	        	curingSize = curingSizeServiceImpl.getCuringSizeById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        curingSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveCuringSize")
	public Response saveCuringSize(final HttpServletRequest req, @RequestBody CuringSize curingSize) throws ResourceNotFoundException {
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
	        	CuringSize savedCuringSize = curingSizeServiceImpl.saveCuringSize(curingSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedCuringSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateCuringSize")
	public Response updateCuringSize(final HttpServletRequest req, @RequestBody CuringSize curingSize) throws ResourceNotFoundException {
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
	        	CuringSize updatedCuringSize = curingSizeServiceImpl.updateCuringSize(curingSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedCuringSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteCuringSize")
	public Response deleteCuringSize(final HttpServletRequest req, @RequestBody CuringSize curingSize) throws ResourceNotFoundException {
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
	        	CuringSize deletedCuringSize = curingSizeServiceImpl.deleteCuringSize(curingSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedCuringSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreCuringSize")
	public Response activateCuringSize(final HttpServletRequest req, @RequestBody CuringSize curingSize) throws ResourceNotFoundException {
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
	        	CuringSize activatedCuringSize = curingSizeServiceImpl.activateCuringSize(curingSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        activatedCuringSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveCuringSizeExcel")
	@Transactional
	public Response saveCuringSizeExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	                List<CuringSize> curingSizes = new ArrayList<>();
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

	                        CuringSize curingSize = new CuringSize();
	                        Cell machineCuringTypeIdCell = row.getCell(2);
	                        Cell sizeIdCell = row.getCell(3);
	                        Cell capacityCell = row.getCell(4);

	                        if (machineCuringTypeIdCell == null || machineCuringTypeIdCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Machine Curing Type ID)");
	                            continue;
	                        }

	                        if (sizeIdCell == null || sizeIdCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Size ID)");
	                            continue;
	                        }

	                        if (capacityCell == null || capacityCell.getCellType() == CellType.BLANK ) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 5 (Capacity)");
	                            continue;
	                        }
	                        
	                        if (capacityCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Tipe Data Salah pada Baris " + (i + 1) + " Kolom 5 (Capacity)");
	                            continue;
	                        }
	                        
	                        Optional<MachineCuringType> machineCuringTypeOpt = machineCuringTypeRepo.findById(machineCuringTypeIdCell.getStringCellValue());
	                        Optional<Size> sizeOpt = sizeRepo.findById(sizeIdCell.getStringCellValue());

	                        if (machineCuringTypeOpt.isPresent()) {
	                            curingSize.setCURINGSIZE_ID(curingSizeServiceImpl.getNewId());
	                            curingSize.setMACHINECURINGTYPE_ID(machineCuringTypeIdCell.getStringCellValue());
	                        } else {
	                            errorMessages.add("Data Tidak Valid, Data Machine Curing Type pada Baris " + (i + 1) + " Tidak Ditemukan");
	                        }

	                        if (sizeOpt.isPresent()) {
	                            if (sizeIdCell != null) {
	                                if (sizeIdCell.getCellType() == CellType.STRING) {
	                                    curingSize.setSIZE_ID(sizeIdCell.getStringCellValue());
	                                } else if (sizeIdCell.getCellType() == CellType.NUMERIC) {
	                                    curingSize.setSIZE_ID(String.valueOf(sizeIdCell.getNumericCellValue()));
	                                }
	                            }
	                        } else {
	                            errorMessages.add("Data Tidak Valid, Data Size pada Baris " + (i + 1) + " Tidak Ditemukan");
	                        }

	                        if (machineCuringTypeOpt.isPresent() && sizeOpt.isPresent()) {
	                            curingSize.setCAPACITY(BigDecimal.valueOf(capacityCell.getNumericCellValue()));
	                            curingSize.setSTATUS(BigDecimal.valueOf(1));
	                            curingSize.setCREATION_DATE(new Date());
	                            curingSize.setLAST_UPDATE_DATE(new Date());

	                            curingSizes.add(curingSize);
	                        }


	                    }
	                }

	                if (!errorMessages.isEmpty()) {
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
	                }

	                curingSizeServiceImpl.deleteAllCuringSize();
	                for (CuringSize curingSize : curingSizes) {
	                    curingSizeServiceImpl.saveCuringSize(curingSize);
	                }

	                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), curingSizes);

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

	
    @GetMapping("/exportCuringSizeexcel")
    public ResponseEntity<InputStreamResource> exportCuringSizesExcel() throws IOException {
        String filename = "EXPORT_MASTER_CURING_SIZE.xlsx";

        ByteArrayInputStream data = curingSizeServiceImpl.exportCuringSizesExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    @GetMapping("/layoutCuringSizeexcel")
    public ResponseEntity<InputStreamResource> layoutCuringSizeexcel() throws IOException {
        String filename = "LAYOUT_MASTER_CURING_SIZE.xlsx";

        ByteArrayInputStream data = curingSizeServiceImpl.layoutCuringSizesExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
}
