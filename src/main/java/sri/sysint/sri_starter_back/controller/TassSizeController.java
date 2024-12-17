package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

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
import sri.sysint.sri_starter_back.model.TassSize;
import sri.sysint.sri_starter_back.repository.MachineTassTypeRepo;
import sri.sysint.sri_starter_back.repository.SizeRepo;
import sri.sysint.sri_starter_back.model.CuringSize;
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.Size;
import sri.sysint.sri_starter_back.service.TassSizeServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class TassSizeController {
	private Response response;	

	@Autowired
	private TassSizeServiceImpl TassSizeServiceImpl;
	
	@Autowired
    private SizeRepo sizeRepo;
	
	@Autowired
    private MachineTassTypeRepo machineTassTypeRepo;

	@PersistenceContext	
	private EntityManager em;
	
	@GetMapping("/getAllTassSize")
	public Response getAllTassSize(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<TassSize> TassSizes = new ArrayList<>();
	    		TassSizes = TassSizeServiceImpl.getAllTassSize();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        TassSizes
			    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	
	@GetMapping("/getTassSizeById/{id}")
	public Response getTassSizeById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<TassSize> tassSize = Optional.of(new TassSize());
	        	tassSize = TassSizeServiceImpl.getTassSizeById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        tassSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveTassSize")
	public Response saveTassSize(final HttpServletRequest req, @RequestBody TassSize tassSize) throws ResourceNotFoundException {
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
	        	TassSize savedTassSize = TassSizeServiceImpl.saveTassSize(tassSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedTassSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateTassSize")
	public Response updateTassSize(final HttpServletRequest req, @RequestBody TassSize tassSize) throws ResourceNotFoundException {
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
	        	TassSize updatedTassSize = TassSizeServiceImpl.updateTassSize(tassSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedTassSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteTassSize")
	public Response deleteTassSize(final HttpServletRequest req, @RequestBody TassSize tassSize) throws ResourceNotFoundException {
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
	        	TassSize deletedTassSize = TassSizeServiceImpl.deleteTassSize(tassSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedTassSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreTassSize")
	public Response activateTassSize(final HttpServletRequest req, @RequestBody TassSize tassSize) throws ResourceNotFoundException {
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
	        	TassSize activatedTassSize = TassSizeServiceImpl.activateTassSize(tassSize);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        activatedTassSize
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveTassSizeExcel")
	public Response saveTassSizeExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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
	    	        
	    	        List<TassSize> TassSizes = new ArrayList<>();
					List<Size> activeSizes = sizeRepo.findSizeActive();
					List<String> errorMessages = new ArrayList<>();
					List<MachineTassType> activeTassTypes = machineTassTypeRepo.findMachineTassTypeActive();
					List<String> sizeIDs = activeSizes.stream()
						.map(Size::getSIZE_ID)
						.collect(Collectors.toList());
					List<String> tassTypeIDs = activeTassTypes.stream()
						.map(MachineTassType::getMACHINETASSTYPE_ID)
						.collect(Collectors.toList());
	    	        
	    	        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);

	                    if (row != null) {
	                        // Check if the row is empty (ignoring borders)
	                        boolean isEmptyRow = true;

	                        for (int j = 0; j < row.getLastCellNum(); j++) {
	                            Cell cell = row.getCell(j);
	                            if (cell != null && cell.getCellType() != CellType.BLANK) {
	                                isEmptyRow = false;
	                                break;
	                            }
	                        }

	                        if (isEmptyRow) {
	                            continue; // Skip the row if it's empty
	                        }

	                        TassSize tassSize = new TassSize();
	                        Cell machineTassTypeCell = row.getCell(2);
	                        Cell sizeIdCell = row.getCell(3);
	                        Cell capacityCell = row.getCell(4);
	                        if (machineTassTypeCell == null || machineTassTypeCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Machine Tass Type)");
	                            continue;
	                        }

	                        if (sizeIdCell == null || sizeIdCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Size ID)");
	                            continue;
	                        }
							if (capacityCell == null || capacityCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Capacity)");
	                            continue;
	                        }
							String sizeID = sizeIdCell.getStringCellValue();
							String machineTypeID = machineTassTypeCell.getStringCellValue();

							Optional<Size> sizeOpt = sizeRepo.findById(sizeID);
							Optional<MachineTassType> machineTassTypeOpt = machineTassTypeRepo.findById(machineTypeID);

	                        if (sizeOpt.isPresent() || machineTassTypeOpt.isPresent() ) {
	                        	tassSize.setTASSIZE_ID(TassSizeServiceImpl.getNewId());
	                        	tassSize.setMACHINETASSTYPE_ID(machineTassTypeCell.getStringCellValue());
	    	                	tassSize.setSIZE_ID(sizeIdCell.getStringCellValue());
	    	                	tassSize.setCAPACITY(BigDecimal.valueOf(capacityCell.getNumericCellValue()));
	    	                	tassSize.setSTATUS(BigDecimal.valueOf(1));
	    	                	tassSize.setCREATION_DATE(new Date());
	    	                	tassSize.setLAST_UPDATE_DATE(new Date());
								TassSizes.add(tassSize);
	                        }else {
	                            errorMessages.add("Data Tidak Valid, Data Tass Size pada Baris " + (i + 1) + " Tidak Ditemukan");
	                        }
	                    }
	                }
					if(!errorMessages.isEmpty()){
						return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
					}

					TassSizeServiceImpl.deleteAllTassSize();
					for(TassSize tassSize: TassSizes){
						TassSizeServiceImpl.saveTassSize(tassSize);
					}
	               
	    	        return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), TassSizes);

	    	    } catch (IOException e) {
	    	        return new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file", req.getRequestURI(), null);
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
	
    @GetMapping("/exportTassSizeexcel")
    public ResponseEntity<InputStreamResource> exportTassSizeExcel() throws IOException {
        String filename = "MASTER_TASS_SIZE.xlsx";

        ByteArrayInputStream data = TassSizeServiceImpl.exportTassSizeExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }

    @GetMapping("/layoutTassSizesExcel")
    public ResponseEntity<InputStreamResource> layoutTassSizesExcel() throws IOException {
        String filename = "LAYOUT_MASTER_TASS_SIZE.xlsx";

        ByteArrayInputStream data = TassSizeServiceImpl.layoutTassSizesExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
}
