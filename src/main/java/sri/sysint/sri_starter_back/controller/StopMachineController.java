package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.TimeZone;

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
import sri.sysint.sri_starter_back.model.StopMachine;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.StopMachineServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class StopMachineController {

    private Response response;

    @Autowired
    private StopMachineServiceImpl stopMachineService;

    @PersistenceContext
    private EntityManager em;

    private String validateTokenAndGetUser(HttpServletRequest req) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            return JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                      .build()
                      .verify(token)
                      .getSubject();
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }
    }

    // GET ALL STOP MACHINES
    @GetMapping("/getAllStopMachines")
    public Response getAllStopMachines(final HttpServletRequest req) throws ResourceNotFoundException {
        String user = validateTokenAndGetUser(req);

        if (user != null) {
            List<StopMachine> stopMachines = stopMachineService.getAllStopMachines();
            response = new Response(
                new Date(),
                HttpStatus.OK.value(),
                null,
                HttpStatus.OK.getReasonPhrase(),
                req.getRequestURI(),
                stopMachines
            );
        } else {
            throw new ResourceNotFoundException("User not found");
        }

        return response;
    }

    // GET STOP MACHINE BY ID
    @GetMapping("/getStopMachineById/{id}")
    public Response getStopMachineById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
        String user = validateTokenAndGetUser(req);

        if (user != null) {
            Optional<StopMachine> stopMachine = stopMachineService.getStopMachineById(id);
            response = new Response(
                new Date(),
                HttpStatus.OK.value(),
                null,
                HttpStatus.OK.getReasonPhrase(),
                req.getRequestURI(),
                stopMachine
            );
        } else {
            throw new ResourceNotFoundException("User not found");
        }

        return response;
    }

    // CREATE NEW STOP MACHINE
    @PostMapping("/saveStopMachine")
    public Response saveStopMachine(final HttpServletRequest req, @RequestBody StopMachine stopMachine) throws ResourceNotFoundException {
        String user = validateTokenAndGetUser(req);

        if (user != null) {
        	stopMachine.setSTART_DATE(adjustDate(stopMachine.getSTART_DATE()));
        	stopMachine.setEND_DATE(adjustDate(stopMachine.getEND_DATE()));

            StopMachine savedStopMachine = stopMachineService.saveStopMachine(stopMachine);
            response = new Response(
                new Date(),
                HttpStatus.OK.value(),
                null,
                HttpStatus.OK.getReasonPhrase(),
                req.getRequestURI(),
                savedStopMachine
            );
        } else {
            throw new ResourceNotFoundException("User not found");
        }

        return response;
    }

	private Date adjustDate(Date originalDate) {
	    if (originalDate != null) {
	        LocalDate localDate = originalDate.toInstant()
	                .atZone(ZoneId.systemDefault())
	                .toLocalDate();

	        LocalDate modifiedDate = localDate.minusDays(0);

	        return Date.from(modifiedDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
	    }
	    return null;  
	}
	
    // UPDATE STOP MACHINE
    @PostMapping("/updateStopMachine")
    public Response updateStopMachine(final HttpServletRequest req, @RequestBody StopMachine stopMachine) throws ResourceNotFoundException {
        String user = validateTokenAndGetUser(req);

        if (user != null) {
        	stopMachine.setSTART_DATE(adjustDate(stopMachine.getSTART_DATE()));
        	stopMachine.setEND_DATE(adjustDate(stopMachine.getEND_DATE()));
        	
            StopMachine updatedStopMachine = stopMachineService.updateStopMachine(stopMachine);
            response = new Response(
                new Date(),
                HttpStatus.OK.value(),
                null,
                HttpStatus.OK.getReasonPhrase(),
                req.getRequestURI(),
                updatedStopMachine
            );
        } else {
            throw new ResourceNotFoundException("User not found");
        }

        return response;
    }

    // DELETE STOP MACHINE
    @PostMapping("/deleteStopMachine")
    public Response deleteStopMachine(final HttpServletRequest req, @RequestBody StopMachine stopMachine) throws ResourceNotFoundException {
        String user = validateTokenAndGetUser(req);

        if (user != null) {
            StopMachine deletedStopMachine = stopMachineService.deleteStopMachine(stopMachine);
            response = new Response(
                new Date(),
                HttpStatus.OK.value(),
                null,
                HttpStatus.OK.getReasonPhrase(),
                req.getRequestURI(),
                deletedStopMachine
            );
        } else {
            throw new ResourceNotFoundException("User not found");
        }

        return response;
    }

    // RESTORE STOP MACHINE
    @PostMapping("/restoreStopMachine")
    public Response restoreStopMachine(final HttpServletRequest req, @RequestBody StopMachine stopMachine) throws ResourceNotFoundException {
        String user = validateTokenAndGetUser(req);

        if (user != null) {
            StopMachine restoredStopMachine = stopMachineService.restoreStopMachine(stopMachine);
            response = new Response(
                new Date(),
                HttpStatus.OK.value(),
                null,
                HttpStatus.OK.getReasonPhrase(),
                req.getRequestURI(),
                restoredStopMachine
            );
        } else {
            throw new ResourceNotFoundException("User not found");
        }

        return response;
    }
    
    @PostMapping("/saveStopMachinesExcel")
    public Response saveStopMachinesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

                stopMachineService.deleteAllStopMachines();

                try (InputStream inputStream = file.getInputStream()) {
                    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                    XSSFSheet sheet = workbook.getSheetAt(0);

                    List<StopMachine> stopMachines = new ArrayList<>();

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

                            // Skip empty rows
                            if (isEmptyRow) {
                                continue;
                            }

                            StopMachine stopMachine = new StopMachine();

                            Cell workCenterCell = row.getCell(2);
                            Cell startDateCell = row.getCell(3);
                            Cell endDateCell = row.getCell(4);

                            if (workCenterCell != null && startDateCell != null && endDateCell != null) {
                                stopMachine.setSTOP_MACHINE_ID(stopMachineService.getNewId());
                                stopMachine.setWORK_CENTER_TEXT(workCenterCell.getStringCellValue());

                                // Convert date values using the helper function for date parsing
                                Date startDate = getDateFromCell(startDateCell);
                                Date endDate = getDateFromCell(endDateCell);

                                if (startDate != null && endDate != null) {
                                    stopMachine.setSTART_DATE(startDate);
                                    stopMachine.setEND_DATE(endDate);

                                    stopMachine.setSTATUS(BigDecimal.valueOf(1));
                                    stopMachine.setCREATION_DATE(new Date());
                                    stopMachine.setLAST_UPDATE_DATE(new Date());
                                }
                            }

                            stopMachineService.saveStopMachine(stopMachine);
                            stopMachines.add(stopMachine);
                        }
                    }

                    return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), stopMachines);

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

    private Date getDateFromCell(Cell cell) throws ParseException {
	    if (cell == null) {
	        return null;
	    }

	    if (cell.getCellType() == CellType.NUMERIC) {
	        if (DateUtil.isCellDateFormatted(cell)) {
	            return cell.getDateCellValue();
	        } else {
	            return null;
	        }
	    } else if (cell.getCellType() == CellType.STRING) {
	        String dateStr = cell.getStringCellValue();
	        if (dateStr == null || dateStr.isEmpty()) {
	            return null;
	        }
	        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
	        sdf.setLenient(false);
	        return sdf.parse(dateStr);
	    }

	    return null;
	}

    @RequestMapping("/exportStopMachinesExcel")
    public ResponseEntity<InputStreamResource> exportStopMachinesExcel() throws IOException {
        String filename = "MASTER_STOP_MACHINE.xlsx";

        ByteArrayInputStream data = stopMachineService.exportStopMachinesExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
            .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
            .body(file);
    }

}