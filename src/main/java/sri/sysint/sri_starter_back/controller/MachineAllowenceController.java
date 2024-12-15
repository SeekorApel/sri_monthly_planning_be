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
import sri.sysint.sri_starter_back.model.MachineAllowence;
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.MachineTass;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.MachineAllowenceServiceImpl;

import sri.sysint.sri_starter_back.repository.MachineCuringRepo;
import sri.sysint.sri_starter_back.repository.MachineTassRepo;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineAllowenceController {
		
	private Response response;	

	@Autowired
	private MachineAllowenceServiceImpl machineAllowenceServiceImpl;

	@Autowired
    private MachineCuringRepo machineCuringRepo;

    @Autowired
    private MachineTassRepo machineTassRepo;
	
	@PersistenceContext	
	private EntityManager em;
	

// START - GET MAPPING
	@GetMapping("/getAllMachineAllowence")
	public Response getAllMachineAllowence(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	// Function goes here
	        	List<MachineAllowence> machineAllowences = new ArrayList<>();
	    	    machineAllowences = machineAllowenceServiceImpl.getAllMachineAllowence();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        machineAllowences
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getMachineAllowenceById/{id}")
	public Response getMachineAllowenceById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<MachineAllowence> machineAllowence = Optional.of(new MachineAllowence());
	    	    machineAllowence = machineAllowenceServiceImpl.getMachineAllowenceById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        machineAllowence
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
// END - GET MAPPING

// START - POST MAPPING
	@PostMapping("/saveMachineAllowence")
	public Response saveMachineAllowence(final HttpServletRequest req, @RequestBody MachineAllowence machineAllowence) throws ResourceNotFoundException {
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
	        	MachineAllowence savedMachineAllowence = machineAllowenceServiceImpl.saveMachineAllowence(machineAllowence);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedMachineAllowence
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateMachineAllowence")
	public Response updateMachineAllowence(final HttpServletRequest req, @RequestBody MachineAllowence machineAllowence) throws ResourceNotFoundException {
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
	        	MachineAllowence updatedMachineAllowence = machineAllowenceServiceImpl.updateMachineAllowence(machineAllowence);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedMachineAllowence
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteMachineAllowence")
	public Response deleteMachineAllowence(final HttpServletRequest req, @RequestBody MachineAllowence machineAllowence) throws ResourceNotFoundException {
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
	        	MachineAllowence deletedMachineAllowence = machineAllowenceServiceImpl.deleteMachineAllowence(machineAllowence);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedMachineAllowence
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveMachineAllowencesExcel")
	public Response saveMachineAllowencesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	                List<MachineAllowence> machineAllowences = new ArrayList<>();
	                List<String> errorMessages = new ArrayList<>();

	                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);
	                    
	                    if (row != null) {
	                        boolean isEmptyRow = true;

	                        for (int j = 1; j < row.getLastCellNum(); j++) {
	                            Cell cell = row.getCell(j);
	                            if (cell != null && cell.getCellType() != CellType.BLANK) {
	                                isEmptyRow = false;
	                                break;
	                            }
	                        }

	                        if (isEmptyRow) {
	                            continue; 
	                        }

	                        MachineAllowence machineAllowence = new MachineAllowence();
							MachineCuring machineCuring = new MachineCuring();
							MachineTass machineTass = new MachineTass();

	                        Cell idMachineCell = row.getCell(2);
	                        Cell personResponsibleCell = row.getCell(3);
	                        Cell shift1Cell = row.getCell(4);
	                        Cell shift2Cell = row.getCell(5);
	                        Cell shift3Cell = row.getCell(6);
	                        Cell shift1FridayCell = row.getCell(7);
	                        Cell totalShift123Cell = row.getCell(8);
	                        if (idMachineCell == null || idMachineCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (ID Machine)");
	                            continue;
	                        }

	                        if (personResponsibleCell == null || personResponsibleCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Person Responsible)");
	                            continue;
	                        }
							if (shift1Cell == null || shift1Cell.getCellType() == CellType.BLANK) {
								errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 5 (Shift 1)");
								continue;
							}

							if (shift2Cell == null || shift2Cell.getCellType() == CellType.BLANK) {
								errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 6 (Shift 2)");
								continue;
							}

							if (shift3Cell == null || shift3Cell.getCellType() == CellType.BLANK) {
								errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 7 (Shift 3)");
								continue;
							}

							if (shift1FridayCell == null || shift1FridayCell.getCellType() == CellType.BLANK) {
								errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 8 (Shift 1 Friday)");
								continue;
							}

							if (totalShift123Cell == null || totalShift123Cell.getCellType() == CellType.BLANK) {
								errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 9 (Total Shift 1, 2, 3)");
								continue;
							}
							String machineid = idMachineCell.getStringCellValue();

	                        Optional<MachineCuring> machineCuringOpt = machineCuringRepo.findById(machineid);
	                        Optional<MachineTass> machineTassOpt = machineTassRepo.findById(machineid);
							if(machineCuringOpt.isPresent()||machineTassOpt.isPresent()){

	                            machineAllowence.setMACHINE_ALLOW_ID(machineAllowenceServiceImpl.getNewId());
	                            machineAllowence.setID_MACHINE(idMachineCell.getStringCellValue());
	                            machineAllowence.setPERSON_RESPONSIBLE(personResponsibleCell.getStringCellValue());
	                            machineAllowence.setSHIFT_1(BigDecimal.valueOf(shift1Cell.getNumericCellValue()));
	                            machineAllowence.setSHIFT_2(BigDecimal.valueOf(shift2Cell.getNumericCellValue()));
	                            machineAllowence.setSHIFT_3(BigDecimal.valueOf(shift3Cell.getNumericCellValue()));
	                            machineAllowence.setSHIFT_1_FRIDAY(BigDecimal.valueOf(shift1FridayCell.getNumericCellValue()));
	                            machineAllowence.setTOTAL_SHIFT_123(BigDecimal.valueOf(totalShift123Cell.getNumericCellValue()));
	                            machineAllowence.setSTATUS(BigDecimal.valueOf(1));
	                            machineAllowence.setCREATION_DATE(new Date());
	                            machineAllowence.setLAST_UPDATE_DATE(new Date());

	                            machineAllowences.add(machineAllowence);
	                        }else{
	                            errorMessages.add("Data Tidak Valid, Data ID Machine pada Baris " + (i + 1) + " Tidak Ditemukan");
							}
	                    }
	                }

					if(!errorMessages.isEmpty()){
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
					}
					machineAllowenceServiceImpl.deleteAllMachineAllowence();
					for(MachineAllowence machineAllowence : machineAllowences){
						machineAllowenceServiceImpl.saveMachineAllowence(machineAllowence);
					}
	                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), machineAllowences);

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
	
	@PostMapping("/restoreMachineAllowence")
	public Response restoreMachineAllowence(final HttpServletRequest req, @RequestBody MachineAllowence machineAllowence) throws ResourceNotFoundException {
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
	        	MachineAllowence deletedMachineAllowence = machineAllowenceServiceImpl.restoreMachineAllowence(machineAllowence);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedMachineAllowence
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
    @GetMapping("/exportMachineAllowenceExcel")
    public ResponseEntity<InputStreamResource> exportMachineAllowenceExcel() throws IOException {
        String filename = "MASTER_MACHINE_ALLOWENCE.xlsx";
        ByteArrayInputStream data = machineAllowenceServiceImpl.exportMachineAllowenceExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }

	@GetMapping("/layoutMachineAllowencesExcel")
    public ResponseEntity<InputStreamResource> layoutQuadrantsExcel() throws IOException {
        String filename = "LAYOUT_MASTER_MACHINE_ALLOWENCE.xlsx";
        ByteArrayInputStream data = machineAllowenceServiceImpl.layoutMachineAllowencesExcel();
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


