package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.security.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
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
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;
import com.auth0.jwt.exceptions.JWTVerificationException;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.model.DWorkDayHours; // Import DWorkDayHours model
import sri.sysint.sri_starter_back.service.DWorkDayHoursServiceImpl; // Import DWorkDayHours service

@CrossOrigin(maxAge = 3600)
@RestController
public class DWorkDayHoursController {

    private Response response;

    @Autowired
    private DWorkDayHoursServiceImpl dWorkDayHoursServiceImpl; 

    @PersistenceContext
    private EntityManager em;

    @GetMapping("/getAllDWorkDayHours")
    public Response getAllDWorkDayHours(final HttpServletRequest req) throws ResourceNotFoundException {
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
                List<DWorkDayHours> dWorkDayHours = dWorkDayHoursServiceImpl.getAllWorkDayHours();

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    dWorkDayHours
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @GetMapping("/getDWorkDayHoursByDate/{date}")
    public Response getDWorkDayHoursByDate(final HttpServletRequest req, @PathVariable String date) throws ResourceNotFoundException {
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
                SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
                Date parsedDate = dateFormat.parse(date); 

                Optional<DWorkDayHours> dWorkDayHours = dWorkDayHoursServiceImpl.getWorkDayHoursByDate(parsedDate); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    dWorkDayHours
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @GetMapping("/getDWorkDayHoursByDateDesc/{date}/{description}")
    public Response getWorkDayByDateDesc(final HttpServletRequest req, 
                                          @PathVariable String date, 
                                          @PathVariable String description) throws ResourceNotFoundException {
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
                SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
                Date parsedDate;
                try {
                    parsedDate = dateFormat.parse(date);
                } catch (ParseException e) {
                    throw new ResourceNotFoundException("Invalid date format, expected format is dd-MM-yyyy");
                }

                Optional<DWorkDayHours> workDayHours = dWorkDayHoursServiceImpl.getWorkDayHoursByDateDesc(parsedDate, description);

              response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        workDayHours
                    );
                
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        } 

        return response;
    }

    @PostMapping("/saveDWorkDayHours")
    public Response saveDWorkDayHours(final HttpServletRequest req, @RequestBody DWorkDayHours dWorkDayHours) throws ResourceNotFoundException {
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
                DWorkDayHours savedDWorkDayHours = dWorkDayHoursServiceImpl.saveWorkDayHours(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    savedDWorkDayHours
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/updateDWorkDayHours")
    public Response updateDWorkDayHours(final HttpServletRequest req, @RequestBody DWorkDayHours dWorkDayHours) throws ResourceNotFoundException {
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
                DWorkDayHours updatedDWorkDayHours = dWorkDayHoursServiceImpl.updateWorkDayHours(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    updatedDWorkDayHours
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/deleteDWorkDayHours")
    public Response deleteDWorkDayHours(final HttpServletRequest req, @RequestBody DWorkDayHours dWorkDayHours) throws ResourceNotFoundException {
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
                DWorkDayHours deletedDWorkDayHours = dWorkDayHoursServiceImpl.deleteWorkDayHours(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    deletedDWorkDayHours
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @PostMapping("/restoreDWorkDayHours")
    public Response restoreDWorkDayHours(final HttpServletRequest req, @RequestBody DWorkDayHours dWorkDayHours) throws ResourceNotFoundException {
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
                DWorkDayHours deletedDWorkDayHours = dWorkDayHoursServiceImpl.restoreWorkDayHours(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    deletedDWorkDayHours
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
//    @PostMapping("/saveDWorkDayHoursExcel")
//    public Response saveDWorkDayHoursExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
//        String header = req.getHeader("Authorization");
//
//        if (header == null || !header.startsWith("Bearer ")) {
//            throw new ResourceNotFoundException("JWT token not found or invalid");
//        }
//
//        String token = header.replace("Bearer ", "");
//
//        try {
//            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
//                .build()
//                .verify(token)
//                .getSubject();
//
//            if (user != null) {
//                if (file.isEmpty()) {
//                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "No file uploaded", req.getRequestURI(), null);
//                }
//
//                dWorkDayHoursServiceImpl.deleteAllWorkDayHours();
//                List<DWorkDayHours> dWorkDayHoursList = new ArrayList<>();
//
//                try (InputStream inputStream = file.getInputStream()) {
//                    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
//                    XSSFSheet sheet = workbook.getSheetAt(0);
//
//                    for (int i = 1; i <= sheet.getLastRowNum(); i++) { 
//                        Row row = sheet.getRow(i);
//
//                        if (row != null) {
//                            DWorkDayHours dWorkDayHours = new DWorkDayHours();
//
//                            Cell idCell = row.getCell(1);
//                            if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
//                                dWorkDayHours.setDETAIL_WD_HOURS_ID(BigDecimal.valueOf(idCell.getNumericCellValue()));
//                            } else {
//                                continue; 
//                            }
//
//                         // Parse DATE_WD (index 2)
//                            Cell dateCell = row.getCell(2);
//                            if (dateCell != null && dateCell.getCellType() == CellType.STRING) {
//                                try {
//                                    // Parse the date using the new formatter to match Excel format with day abbreviation
//                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEE MMMM dd, yyyy", Locale.ENGLISH);
//                                    LocalDate workDate = LocalDate.parse(dateCell.getStringCellValue(), formatter);
//
//                                    // Set DATE_WD explicitly to midnight UTC
//                                    ZonedDateTime midnightUtc = workDate.atStartOfDay(ZoneId.of("UTC"));
//                                    Date midnightDate = Date.from(midnightUtc.toInstant()); // Convert to java.util.Date
//
//                                    dWorkDayHours.setDATE_WD(midnightDate); // Store as midnight UTC
//                                } catch (Exception e) {
//                                    System.out.println("Format tanggal tidak sesuai untuk baris ke-" + (i + 1) + ", data: " + dateCell.getStringCellValue());
//                                    continue; // Skip this row if date parsing fails
//                                }
//                            } else {
//                                System.out.println("Cell tanggal kosong atau tidak valid pada baris ke-" + (i + 1));
//                                continue; // Skip this row if the date cell is invalid
//                            }
//
//
//
//                            // Populate HOURS1 to HOURS24 columns (index 3 to 26)
//                            for (int j = 3; j <= 26; j++) {
//                                Cell hourCell = row.getCell(j);
//                                BigDecimal hourValue = (hourCell != null && hourCell.getCellType() == CellType.STRING 
//                                    && "v".equalsIgnoreCase(hourCell.getStringCellValue())) ? BigDecimal.ONE : BigDecimal.ZERO;
//
//                                switch (j) {
//                                    case 3: dWorkDayHours.setHOUR_1(hourValue); break;
//                                    case 4: dWorkDayHours.setHOUR_2(hourValue); break;
//                                    case 5: dWorkDayHours.setHOUR_3(hourValue); break;
//                                    case 6: dWorkDayHours.setHOUR_4(hourValue); break;
//                                    case 7: dWorkDayHours.setHOUR_5(hourValue); break;
//                                    case 8: dWorkDayHours.setHOUR_6(hourValue); break;
//                                    case 9: dWorkDayHours.setHOUR_7(hourValue); break;
//                                    case 10: dWorkDayHours.setHOUR_8(hourValue); break;
//                                    case 11: dWorkDayHours.setHOUR_9(hourValue); break;
//                                    case 12: dWorkDayHours.setHOUR_10(hourValue); break;
//                                    case 13: dWorkDayHours.setHOUR_11(hourValue); break;
//                                    case 14: dWorkDayHours.setHOUR_12(hourValue); break;
//                                    case 15: dWorkDayHours.setHOUR_13(hourValue); break;
//                                    case 16: dWorkDayHours.setHOUR_14(hourValue); break;
//                                    case 17: dWorkDayHours.setHOUR_15(hourValue); break;
//                                    case 18: dWorkDayHours.setHOUR_16(hourValue); break;
//                                    case 19: dWorkDayHours.setHOUR_17(hourValue); break;
//                                    case 20: dWorkDayHours.setHOUR_18(hourValue); break;
//                                    case 21: dWorkDayHours.setHOUR_19(hourValue); break;
//                                    case 22: dWorkDayHours.setHOUR_20(hourValue); break;
//                                    case 23: dWorkDayHours.setHOUR_21(hourValue); break;
//                                    case 24: dWorkDayHours.setHOUR_22(hourValue); break;
//                                    case 25: dWorkDayHours.setHOUR_23(hourValue); break;
//                                    case 26: dWorkDayHours.setHOUR_24(hourValue); break;
//                                    default: break;
//                                }
//                            }
//
//                            // Parse DESCRIPTION (index 27)
//                            Cell descriptionCell = row.getCell(27);
//                            if (descriptionCell != null && descriptionCell.getCellType() == CellType.STRING) {
//                                dWorkDayHours.setDESCRIPTION(descriptionCell.getStringCellValue());
//                            }
//
//                            // Set additional fields for DWorkDayHours
//                            dWorkDayHours.setSTATUS(BigDecimal.valueOf(1)); // Active status
//                            dWorkDayHours.setCREATION_DATE(new Date());
//                            dWorkDayHours.setLAST_UPDATE_DATE(new Date());
//
//                            // Save DWorkDayHours to the database
//                            dWorkDayHoursServiceImpl.saveWorkDayHours(dWorkDayHours);
//                            dWorkDayHoursList.add(dWorkDayHours);
//                        }
//                    }
//
//                    return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), dWorkDayHoursList);
//
//                } catch (IOException e) {
//                    return new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file: " + e.getMessage(), req.getRequestURI(), null);
//                }
//            } else {
//                throw new ResourceNotFoundException("User not found");
//            }
//        } catch (Exception e) {
//            throw new ResourceNotFoundException("JWT token is not valid or expired");
//        }
//    }
    
    @PostMapping("/saveDWorkDayHoursExcel")
    public Response saveDWorkDayHoursExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or invalid");
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

                dWorkDayHoursServiceImpl.deleteAllWorkDayHours();
                List<DWorkDayHours> dWorkDayHoursList = new ArrayList<>();

                try (InputStream inputStream = file.getInputStream()) {
                    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                    XSSFSheet sheet = workbook.getSheetAt(0);

                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);

                        if (row != null) {
                            DWorkDayHours dWorkDayHours = new DWorkDayHours();

                            Cell idCell = row.getCell(1);
                            if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                                dWorkDayHours.setDETAIL_WD_HOURS_ID(BigDecimal.valueOf(idCell.getNumericCellValue()));
                            } else {
                                continue; 
                            }

                            Cell dateCell = row.getCell(2);
                            if (dateCell != null) {
                                try {
                                    LocalDate workDate;
                                    if (dateCell.getCellType() == CellType.NUMERIC) {
                                        workDate = dateCell.getLocalDateTimeCellValue().toLocalDate();
                                    } else if (dateCell.getCellType() == CellType.STRING) {
                                        workDate = LocalDate.parse(dateCell.getStringCellValue(), DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                                    } else {
                                        System.out.println("Cell tanggal tidak valid pada baris ke-" + (i + 1));
                                        continue;
                                    }

                                    ZonedDateTime midnightUtc = workDate.atStartOfDay(ZoneId.of("UTC"));
                                    Date midnightDate = Date.from(midnightUtc.toInstant()); 
                                    dWorkDayHours.setDATE_WD(midnightDate);

                                    System.out.println("Tanggal berhasil di-parse dan diatur ke UTC tengah malam: " + midnightUtc);
                                } catch (Exception e) {
                                    System.out.println("Format tanggal tidak sesuai untuk baris ke-" + (i + 1) + ", data: " + dateCell.getStringCellValue());
                                    continue;
                                }
                            } else {
                                System.out.println("Cell tanggal kosong atau tidak valid pada baris ke-" + (i + 1));
                                continue;
                            }

                            for (int j = 3; j <= 26; j++) {
                                Cell hourCell = row.getCell(j);
                                BigDecimal hourValue = (hourCell != null && hourCell.getCellType() == CellType.STRING 
                                    && "v".equalsIgnoreCase(hourCell.getStringCellValue())) ? BigDecimal.ONE : BigDecimal.ZERO;

                                switch (j) {
                                    case 3: dWorkDayHours.setHOUR_1(hourValue); break;
                                    case 4: dWorkDayHours.setHOUR_2(hourValue); break;
                                    case 5: dWorkDayHours.setHOUR_3(hourValue); break;
                                    case 6: dWorkDayHours.setHOUR_4(hourValue); break;
                                    case 7: dWorkDayHours.setHOUR_5(hourValue); break;
                                    case 8: dWorkDayHours.setHOUR_6(hourValue); break;
                                    case 9: dWorkDayHours.setHOUR_7(hourValue); break;
                                    case 10: dWorkDayHours.setHOUR_8(hourValue); break;
                                    case 11: dWorkDayHours.setHOUR_9(hourValue); break;
                                    case 12: dWorkDayHours.setHOUR_10(hourValue); break;
                                    case 13: dWorkDayHours.setHOUR_11(hourValue); break;
                                    case 14: dWorkDayHours.setHOUR_12(hourValue); break;
                                    case 15: dWorkDayHours.setHOUR_13(hourValue); break;
                                    case 16: dWorkDayHours.setHOUR_14(hourValue); break;
                                    case 17: dWorkDayHours.setHOUR_15(hourValue); break;
                                    case 18: dWorkDayHours.setHOUR_16(hourValue); break;
                                    case 19: dWorkDayHours.setHOUR_17(hourValue); break;
                                    case 20: dWorkDayHours.setHOUR_18(hourValue); break;
                                    case 21: dWorkDayHours.setHOUR_19(hourValue); break;
                                    case 22: dWorkDayHours.setHOUR_20(hourValue); break;
                                    case 23: dWorkDayHours.setHOUR_21(hourValue); break;
                                    case 24: dWorkDayHours.setHOUR_22(hourValue); break;
                                    case 25: dWorkDayHours.setHOUR_23(hourValue); break;
                                    case 26: dWorkDayHours.setHOUR_24(hourValue); break;
                                    default: break;
                                }
                            }

                            Cell descriptionCell = row.getCell(27);
                            if (descriptionCell != null && descriptionCell.getCellType() == CellType.STRING) {
                                dWorkDayHours.setDESCRIPTION(descriptionCell.getStringCellValue());
                            }

                            dWorkDayHours.setSTATUS(BigDecimal.valueOf(1)); // Active status
                            dWorkDayHours.setCREATION_DATE(new Date());
                            dWorkDayHours.setLAST_UPDATE_DATE(new Date());

                            dWorkDayHoursServiceImpl.saveWorkDayHours(dWorkDayHours);
                            dWorkDayHoursList.add(dWorkDayHours);
                        }
                    }

                    return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), dWorkDayHoursList);

                } catch (IOException e) {
                    return new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file: " + e.getMessage(), req.getRequestURI(), null);
                }
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }
    }
    
    @RequestMapping("/exportDWorkDayHoursExcel")
    public ResponseEntity<InputStreamResource> exportDWorkDayHoursExcel() throws IOException {
        String filename = "MASTER_WORKDAY_HOURS.xlsx";

        ByteArrayInputStream data = dWorkDayHoursServiceImpl.exportWorkDayHoursExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    @PostMapping("/turnOnHour/{dateWd}/{hour}/{description}")
    public Response turnOnHour(final HttpServletRequest req, @PathVariable String dateWd, @PathVariable String hour,@PathVariable String description) throws ResourceNotFoundException {
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
            	DWorkDayHours updatedWorkDay = dWorkDayHoursServiceImpl.turnOnHour(dateWd, hour, description);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    updatedWorkDay
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/turnOffHour/{dateWd}/{hour}/{description}")
    public Response turnOffHour(final HttpServletRequest req, @PathVariable String dateWd, @PathVariable String hour, @PathVariable String description) throws ResourceNotFoundException {
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
            	DWorkDayHours updatedWorkDay = dWorkDayHoursServiceImpl.turnOffHour(dateWd, hour, description);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    updatedWorkDay
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
}
