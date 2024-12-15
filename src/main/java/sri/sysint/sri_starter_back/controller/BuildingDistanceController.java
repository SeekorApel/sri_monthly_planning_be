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
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.Building;
import sri.sysint.sri_starter_back.model.BuildingDistance;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.BuildingRepo;
import sri.sysint.sri_starter_back.service.BuildingDistanceServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class BuildingDistanceController {
		
	private Response response;	

	@Autowired
	private BuildingDistanceServiceImpl buildingDistanceServiceImpl;
	@Autowired
    private BuildingRepo buildingRepo;
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/getAllBuildingDistance")
	public Response getAllBuildingDistance(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<BuildingDistance> buildingDistances = new ArrayList<>();
	    	    buildingDistances = buildingDistanceServiceImpl.getAllBuildingDistance();

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        buildingDistances
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getBuildingDistanceById/{id}")
	public Response getBuildingDistanceById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<BuildingDistance> buildingDistance = Optional.of(new BuildingDistance());
	    	    buildingDistance = buildingDistanceServiceImpl.getBuildingDistanceById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        buildingDistance
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
	@PostMapping("/saveBuildingDistance")
	public Response saveBuildingDistance(final HttpServletRequest req, @RequestBody BuildingDistance buildingDistance) throws ResourceNotFoundException {
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
	        	BuildingDistance savedBuildingDistance = buildingDistanceServiceImpl.saveBuildingDistance(buildingDistance);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedBuildingDistance
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateBuildingDistance")
	public Response updateBuildingDistance(final HttpServletRequest req, @RequestBody BuildingDistance buildingDistance) throws ResourceNotFoundException {
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
	        	BuildingDistance updatedBuildingDistance = buildingDistanceServiceImpl.updateBuildingDistance(buildingDistance);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedBuildingDistance
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteBuildingDistance")
	public Response deleteBuildingDistance(final HttpServletRequest req, @RequestBody BuildingDistance buildingDistance) throws ResourceNotFoundException {
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
	        	BuildingDistance deletedBuildingDistance = buildingDistanceServiceImpl.deleteBuildingDistance(buildingDistance);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedBuildingDistance
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreBuildingDistance")
	public Response restoreBuildingDistance(final HttpServletRequest req, @RequestBody BuildingDistance buildingDistance) throws ResourceNotFoundException {
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
	            BuildingDistance restoredBuildingDistance = buildingDistanceServiceImpl.restoreBuildingDistance(buildingDistance);

	            response = new Response(
	                new Date(),
	                HttpStatus.OK.value(),
	                null,
	                HttpStatus.OK.getReasonPhrase(),
	                req.getRequestURI(),
	                restoredBuildingDistance
	            );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}

	
	@PostMapping("/saveBuildingDistancesExcel")
	@Transactional
	public Response saveBuildingDistancesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

	                List<BuildingDistance> buildingDistances = new ArrayList<>();
	                List<String> errorMessages = new ArrayList<>();

	                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	                    Row row = sheet.getRow(i);

	                    if (row != null) {
	                        boolean isEmptyRow = true;

	                        // Check if the row is empty
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

	                        BuildingDistance buildingDistance = new BuildingDistance();

	                        // Read BUILDING_NAME_1, BUILDING_NAME_2, and DISTANCE from the Excel file
	                        Cell buildingName1Cell = row.getCell(2);
	                        Cell buildingName2Cell = row.getCell(3);
	                        Cell distanceCell = row.getCell(4);

	                        if (buildingName1Cell == null || buildingName1Cell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (Building Name 1)");
	                            continue;
	                        }

	                        if (buildingName2Cell == null || buildingName2Cell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Building Name 2)");
	                            continue;
	                        }

	                        if (distanceCell == null || distanceCell.getCellType() != CellType.NUMERIC) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong atau Tidak Valid pada Baris " + (i + 1) + " Kolom 5 (Distance)");
	                            continue;
	                        }

	                        // Find the BUILDING_ID by name for BUILDING_NAME_1
	                        Optional<Building> building1Opt = buildingRepo.findByName(buildingName1Cell.getStringCellValue());
	                        if (building1Opt.isPresent()) {
	                            buildingDistance.setBUILDING_ID_1(building1Opt.get().getBUILDING_ID());
	                        } else {
	                            errorMessages.add("Data Tidak Valid, BUILDING_NAME_1 pada Baris " + (i + 1) + " Tidak Ditemukan");
	                            continue;
	                        }

	                        // Find the BUILDING_ID by name for BUILDING_NAME_2
	                        Optional<Building> building2Opt = buildingRepo.findByName(buildingName2Cell.getStringCellValue());
	                        if (building2Opt.isPresent()) {
	                            buildingDistance.setBUILDING_ID_2(building2Opt.get().getBUILDING_ID());
	                        } else {
	                            errorMessages.add("Data Tidak Valid, BUILDING_NAME_2 pada Baris " + (i + 1) + " Tidak Ditemukan");
	                            continue;
	                        }

	                        // Set other fields
	                        buildingDistance.setID_B_DISTANCE(buildingDistanceServiceImpl.getNewId());
	                        buildingDistance.setDISTANCE(BigDecimal.valueOf(distanceCell.getNumericCellValue()));
	                        buildingDistance.setSTATUS(BigDecimal.valueOf(1));
	                        buildingDistance.setCREATION_DATE(new Date());
	                        buildingDistance.setLAST_UPDATE_DATE(new Date());

	                        buildingDistances.add(buildingDistance);
	                    }
	                }

	                if (!errorMessages.isEmpty()) {
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
	                }

	                // Delete all previous building distances before saving the new ones
	                buildingDistanceServiceImpl.deleteAllBuildingDistance();
	                for (BuildingDistance buildingDistance : buildingDistances) {
	                    buildingDistanceServiceImpl.saveBuildingDistance(buildingDistance);
	                }

	                return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), buildingDistances);

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

	
    @RequestMapping("/exportBuildingDistancesExcel")
    public ResponseEntity<InputStreamResource> exportBuildingDistancesExcel() throws IOException {
        String filename = "EXPORT_MASTER_BUILDING_DISTANCE.xlsx";

        ByteArrayInputStream data = buildingDistanceServiceImpl.exportBuildingDistancesExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }

    @RequestMapping("/layoutBuildingDistancesExcel")
    public ResponseEntity<InputStreamResource> layoutBuildingDistancesExcel() throws IOException {
        String filename = "LAYOUT_MASTER_BUILDING_DISTANCE.xlsx";

        ByteArrayInputStream data = buildingDistanceServiceImpl.layoutBuildingDistancesExcel();
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
