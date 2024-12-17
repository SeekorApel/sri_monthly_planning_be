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
import sri.sysint.sri_starter_back.model.MachineCuringType;
import sri.sysint.sri_starter_back.model.MachineTassType;
import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.Setting;
import sri.sysint.sri_starter_back.repository.SettingRepo;
import sri.sysint.sri_starter_back.service.MachineCuringTypeServiceImpl;
import sri.sysint.sri_starter_back.service.MachineTassTypeServiceImpl;
import sri.sysint.sri_starter_back.service.PlantServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineTassTypeController {
		
	private Response response;	

	@Autowired
	private MachineTassTypeServiceImpl machineTassTypeServiceImpl;
	@Autowired
	private SettingRepo settingRepo;
	@PersistenceContext	
	private EntityManager em;
	
	//START - GET MAPPING
		@GetMapping("/getAllMachineTassType")
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
		        	List<MachineTassType> machineTassTypes = new ArrayList<>();
		        	machineTassTypes = machineTassTypeServiceImpl.getAllMachineTassType();

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        machineTassTypes
		    	    );
		        } else {
		            throw new ResourceNotFoundException("User not found");
		        }
		    } catch (Exception e) {
		        throw new ResourceNotFoundException("JWT token is not valid or expired");
		    }

		    return response;
		}
	
		@GetMapping("/getMachineTassTypeById/{id}")
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
		        	Optional<MachineTassType> machineTassType = Optional.of(new MachineTassType());
		        	machineTassType = machineTassTypeServiceImpl.getMachineTassTypeById(id);

		    	    response = new Response(
		    	        new Date(),
		    	        HttpStatus.OK.value(),
		    	        null,
		    	        HttpStatus.OK.getReasonPhrase(),
		    	        req.getRequestURI(),
		    	        machineTassType
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
			@PostMapping("/saveMachineTassType")
			public Response savePlant(final HttpServletRequest req, @RequestBody MachineTassType machineTassType) throws ResourceNotFoundException {
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
			        	MachineTassType savedMachineTassType = machineTassTypeServiceImpl.saveMachineTassType(machineTassType);

			    	    response = new Response(
			    	        new Date(),
			    	        HttpStatus.OK.value(),
			    	        null,
			    	        HttpStatus.OK.getReasonPhrase(),
			    	        req.getRequestURI(),
			    	        savedMachineTassType
			    	    );
			        } else {
			            throw new ResourceNotFoundException("User not found");
			        }
			    } catch (Exception e) {
			        throw new ResourceNotFoundException("JWT token is not valid or expired");
			    }

			    return response;
			}
			
			@PostMapping("/updateMachineTassType")
			public Response updatePlant(final HttpServletRequest req, @RequestBody MachineTassType machineTassType) throws ResourceNotFoundException {
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
			        	MachineTassType updatedMachineTassType = machineTassTypeServiceImpl.updateMachineTassType(machineTassType);

			    	    response = new Response(
			    	        new Date(),
			    	        HttpStatus.OK.value(),
			    	        null,
			    	        HttpStatus.OK.getReasonPhrase(),
			    	        req.getRequestURI(),
			    	        updatedMachineTassType
			    	    );
			        } else {
			            throw new ResourceNotFoundException("User not found");
			        }
			    } catch (Exception e) {
			        throw new ResourceNotFoundException("JWT token is not valid or expired");
			    }

			    return response;
			}
			
			@PostMapping("/deleteMachineTassType")
			public Response deletetePlant(final HttpServletRequest req, @RequestBody MachineTassType machineTassType) throws ResourceNotFoundException {
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
			        	MachineTassType deletedMachineTassType = machineTassTypeServiceImpl.deleteMachineTassType(machineTassType);

			    	    response = new Response(
			    	        new Date(),
			    	        HttpStatus.OK.value(),
			    	        null,
			    	        HttpStatus.OK.getReasonPhrase(),
			    	        req.getRequestURI(),
			    	        deletedMachineTassType
			    	    );
			        } else {
			            throw new ResourceNotFoundException("User not found");
			        }
			    } catch (Exception e) {
			        throw new ResourceNotFoundException("JWT token is not valid or expired");
			    }

			    return response;
			}
			
			@PostMapping("/restoreMachineTassType")
			public Response restoreMachineTassType(final HttpServletRequest req, @RequestBody MachineTassType machineTassType) throws ResourceNotFoundException {
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
			            MachineTassType restoredMachineTassType = machineTassTypeServiceImpl.restoreMachineTassType(machineTassType);

			            response = new Response(
			                new Date(),
			                HttpStatus.OK.value(),
			                null,
			                HttpStatus.OK.getReasonPhrase(),
			                req.getRequestURI(),
			                restoredMachineTassType
			            );
			        } else {
			            throw new ResourceNotFoundException("User not found");
			        }
			    } catch (Exception e) {
			        throw new ResourceNotFoundException("JWT token is not valid or expired");
			    }

			    return response;
			}
			
			@PostMapping("/saveMachineTassTypeExcel")
			@Transactional
			public Response saveMachineTassTypesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

			                List<MachineTassType> machineTassTypes = new ArrayList<>();
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

			                        MachineTassType machineTassType = new MachineTassType();
			                        Cell codeCell = row.getCell(1);
			                        Cell settingValueCell = row.getCell(2);
			                        Cell descriptionCell = row.getCell(3);

			                        if (codeCell == null || codeCell.getCellType() == CellType.BLANK) {
			                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 2 (Machine Tass Type ID)");
			                            continue;
			                        }

			                        if (settingValueCell == null || settingValueCell.getCellType() == CellType.BLANK) {
			                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Setting Value)");
			                            continue;
			                        }

			                        if (descriptionCell == null || descriptionCell.getCellType() == CellType.BLANK) {
			                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Description)");
			                            continue;
			                        }

			                        String settingValue = null;
			                        if (settingValueCell.getCellType() == CellType.NUMERIC) {
			                            settingValue = String.valueOf((int) settingValueCell.getNumericCellValue());
			                        } else if (settingValueCell.getCellType() == CellType.STRING) {
			                            settingValue = settingValueCell.getStringCellValue();
			                        }

			                        Optional<Setting> settingOptional = settingRepo.findBySettingValue(settingValue);
			                        if (settingOptional.isPresent()) {
			                        	machineTassType.setSETTING_ID(settingOptional.get().getSETTING_ID());
			                        } else {
			                            errorMessages.add("Data Tidak Valid, Data Setting pada Baris " + (i + 1) + " Tidak Ditemukan");
			                            continue;
			                        }
			                            machineTassType.setMACHINETASSTYPE_ID(codeCell.getStringCellValue());
			                            machineTassType.setSETTING_ID(settingOptional.get().getSETTING_ID());
			                            machineTassType.setDESCRIPTION(descriptionCell.getStringCellValue());
			                            machineTassType.setSTATUS(BigDecimal.valueOf(1));
			                            machineTassType.setCREATION_DATE(new Date());
			                            machineTassType.setLAST_UPDATE_DATE(new Date());
			                            
			                            machineTassTypes.add(machineTassType);

			                    }
			                }

			                if (!errorMessages.isEmpty()) {
			                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
			                }

			                machineTassTypeServiceImpl.deleteAllMachineTassType();
			                for (MachineTassType machineTassType : machineTassTypes) {
			                    machineTassTypeServiceImpl.saveMachineTassType(machineTassType);
			                }

			                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), machineTassTypes);

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

		    @GetMapping("/exportMachineTassTypeExcel")
		    public ResponseEntity<InputStreamResource> exportMachineTassTypesExcel() throws IOException {
		        String filename = "EXPORT_MACHINE_TASS_TYPE.xlsx";

		        ByteArrayInputStream data = machineTassTypeServiceImpl.exportMachineTassTypesExcel();
		        InputStreamResource file = new InputStreamResource(data);

		        return ResponseEntity.ok()
		                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
		                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
		                .body(file);
		    }
			
		    @GetMapping("/layoutMachineTassTypeExcel")
		    public ResponseEntity<InputStreamResource> layoutMachineTassTypeExcel() throws IOException {
		        String filename = "LAYOUT_MACHINE_TASS_TYPE.xlsx";

		        ByteArrayInputStream data = machineTassTypeServiceImpl.layoutMachineTassTypesExcel();
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


