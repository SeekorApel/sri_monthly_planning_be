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
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
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
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.DDeliverySchedule;
import sri.sysint.sri_starter_back.model.DeliverySchedule;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.repository.DeliveryScheduleRepo;
import sri.sysint.sri_starter_back.service.DDeliveryScheduleServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class DDeliveryScheduleController {
    
    private Response response;    

    @Autowired
    private DDeliveryScheduleServiceImpl dDeliveryScheduleServiceImpl;
    @Autowired
    private DeliveryScheduleRepo deliveryScheduleRepo;

    @PersistenceContext    
    private EntityManager em;

    // START - GET MAPPING
    @GetMapping("/getAllDDeliverySchedule")
    public Response getAllDDeliverySchedule(final HttpServletRequest req) throws ResourceNotFoundException {
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
                List<DDeliverySchedule> dDeliverySchedules = new ArrayList<>();
                dDeliverySchedules = dDeliveryScheduleServiceImpl.getAllDDeliverySchedule();

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    dDeliverySchedules
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @GetMapping("/getDDeliveryScheduleById/{id}")
    public Response getDDeliveryScheduleById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
                Optional<DDeliverySchedule> dDeliverySchedule = Optional.of(new DDeliverySchedule());
                dDeliverySchedule = dDeliveryScheduleServiceImpl.getDDeliveryScheduleById(id);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    dDeliverySchedule
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
    @PostMapping("/saveDDeliverySchedule")
    public Response saveDDeliverySchedule(final HttpServletRequest req, @RequestBody DDeliverySchedule dDeliverySchedule) throws ResourceNotFoundException {
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
            	dDeliverySchedule.setDATE_DS(adjustDate(dDeliverySchedule.getDATE_DS()));
            	
                DDeliverySchedule savedDDeliverySchedule = dDeliveryScheduleServiceImpl.saveDDeliverySchedule(dDeliverySchedule);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    savedDDeliverySchedule
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
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
    
    @PostMapping("/updateDDeliverySchedule")
    public Response updateDDeliverySchedule(final HttpServletRequest req, @RequestBody DDeliverySchedule dDeliverySchedule) throws ResourceNotFoundException {
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
            	dDeliverySchedule.setDATE_DS(adjustDate(dDeliverySchedule.getDATE_DS()));

                DDeliverySchedule updatedDDeliverySchedule = dDeliveryScheduleServiceImpl.updateDDeliverySchedule(dDeliverySchedule);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    updatedDDeliverySchedule
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/deleteDDeliverySchedule")
    public Response deleteDDeliverySchedule(final HttpServletRequest req, @RequestBody DDeliverySchedule dDeliverySchedule) throws ResourceNotFoundException {
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
                DDeliverySchedule deletedDDeliverySchedule = dDeliveryScheduleServiceImpl.deleteDDeliverySchedule(dDeliverySchedule);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    deletedDDeliverySchedule
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @PostMapping("/restoreDDeliverySchedule")
    public Response restoreDDeliverySchedule(final HttpServletRequest req, @RequestBody DDeliverySchedule dDeliverySchedule) throws ResourceNotFoundException {
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
                // Assuming you have a method to restore a DDeliverySchedule
                DDeliverySchedule restoredDDeliverySchedule = dDeliveryScheduleServiceImpl.restoreDDeliverySchedule(dDeliverySchedule);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    restoredDDeliverySchedule
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/saveDDeliverySchedulesExcel")
    public Response saveDDeliverySchedulesExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

                Response response;

                try (InputStream inputStream = file.getInputStream();
                     XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {

                    XSSFSheet sheet = workbook.getSheetAt(0);
                    List<DDeliverySchedule> dDeliverySchedules = new ArrayList<>();
	                List<String> errorMessages = new ArrayList<>();

                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row != null) {
                            DDeliverySchedule dDeliverySchedule = new DDeliverySchedule();

                            Cell dsIdCell = row.getCell(2);
                            Cell partNumCell = row.getCell(3);
                            Cell dateDsCell = row.getCell(4);
                            Cell totalDeliveryCell = row.getCell(5);

	                        if (dsIdCell == null || dsIdCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 3 (DetailDelivery Schedule ID)");
	                            continue;
	                        }

	                        if (partNumCell == null || partNumCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 4 (Partnum)");
	                            continue;
	                        }
                            if (dateDsCell == null || dateDsCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 5 (Date)");
	                            continue;
	                        }

	                        if (totalDeliveryCell == null || totalDeliveryCell.getCellType() == CellType.BLANK) {
	                            errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom 6 (Total Delivery)");
	                            continue;
	                        }
	                        BigDecimal dsIdC;
	                        try {
	                            dsIdC = new BigDecimal(dsIdCell.toString());
	                        } catch (NumberFormatException e) {
	                            errorMessages.add(String.format("Row %d, Column 3 (DetailDelivery Schedule ID) has invalid format", i + 1));
	                            continue;
	                        }
	                        Optional<DeliverySchedule> deliveryScheduleOpt = deliveryScheduleRepo.findById(dsIdC);
                            if (deliveryScheduleOpt.isPresent()) {
                                dDeliverySchedule.setDETAIL_DS_ID(dDeliveryScheduleServiceImpl.getNewId());

                                dDeliverySchedule.setDS_ID(dsIdC);

                                BigDecimal partNum = BigDecimal.valueOf(partNumCell.getNumericCellValue());
                                dDeliverySchedule.setPART_NUM(partNum);

                                try {
                                    Date dateDs;
                                    if (dateDsCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateDsCell)) {
                                        dateDs = dateDsCell.getDateCellValue();
                                    } else if (dateDsCell.getCellType() == CellType.STRING) {
                                        // Parse string date if formatted as text
                                        String dateString = dateDsCell.getStringCellValue();
                                        LocalDate localDate = LocalDate.parse(dateString, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.S"));
                                        dateDs = Date.from(localDate.atStartOfDay(ZoneId.of("UTC")).toInstant());
                                    } else {
                                        continue; // Skip if date format is unrecognized
                                    }

                                    dDeliverySchedule.setDATE_DS(dateDs);

                                } catch (Exception e) {
                                    e.printStackTrace();
                                    continue;
                                }

                                BigDecimal totalDelivery = BigDecimal.valueOf(totalDeliveryCell.getNumericCellValue());
                                dDeliverySchedule.setTOTAL_DELIVERY(totalDelivery);

                                dDeliverySchedule.setSTATUS(BigDecimal.valueOf(1));
                                dDeliverySchedule.setCREATION_DATE(new Date());
                                dDeliverySchedule.setLAST_UPDATE_DATE(new Date());
                                dDeliverySchedules.add(dDeliverySchedule);
                            } else {
	                            errorMessages.add("Data Tidak Valid, Data Delivery Schedule pada Baris " + (i + 1) + " Tidak Ditemukan");
	                        }
                        }
                    }
                    if (!errorMessages.isEmpty()) {
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
	                }
                    dDeliveryScheduleServiceImpl.deleteAllDDeliverySchedule();
                    for(DDeliverySchedule dDeliverySchedule : dDeliverySchedules){
                        dDeliveryScheduleServiceImpl.saveDDeliverySchedule(dDeliverySchedule);
                    }

                    return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), dDeliverySchedules);

                } catch (IOException e) {
                    e.printStackTrace();
                    return new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file: " + e.getMessage(), req.getRequestURI(), null);
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
    
    @GetMapping("/exportDDeliveryScheduleExcel")
    public ResponseEntity<InputStreamResource> exportDDeliveryScheduleExcel() throws IOException {
        String filename = "MASTER_DETAIL_DELIVERY_SCHEDULE.xlsx";

        ByteArrayInputStream data = dDeliveryScheduleServiceImpl.exportDDeliverySchedulesExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }

    @GetMapping("/layoutDDeliveryScheduleExcel")
    public ResponseEntity<InputStreamResource> layoutDDeliveryScheduleExcel() throws IOException {
        String filename = "LAYOUT_MASTER_DELIVERY_SCHEDULE.xlsx";
        ByteArrayInputStream data = dDeliveryScheduleServiceImpl.layoutDDeliveryScheduleExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    // END - POST MAPPING
    // START - PUT MAPPING
    // END - PUT MAPPING
    // START - DELETE MAPPING
    // END - DELETE MAPPING
    // START - PROCEDURE
    // END - PROCEDURE
}
