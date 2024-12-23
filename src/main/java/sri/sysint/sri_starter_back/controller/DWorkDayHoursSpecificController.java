package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.security.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
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
import sri.sysint.sri_starter_back.model.DWorkDayHoursSpesific;
import sri.sysint.sri_starter_back.service.DWorkDayHoursSpecificServiceImpl; 

@CrossOrigin(maxAge = 3600)
@RestController
public class DWorkDayHoursSpecificController {

    private Response response;

    @Autowired
    private DWorkDayHoursSpecificServiceImpl dWorkDayHoursSpecificServiceImpl; 

    @PersistenceContext
    private EntityManager em;

    @GetMapping("/getAllDWorkDayHoursSpecific")
    public Response getAllDWorkDayHoursSpecific(final HttpServletRequest req) throws ResourceNotFoundException {
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
            	List<DWorkDayHoursSpesific> dWorkDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.getAllWorkDayHoursSpecific();
            	System.out.println("Data fetched: " + dWorkDayHoursSpecific);
            	
            	response = new Response(
            	    new Date(),
            	    HttpStatus.OK.value(),
            	    null,
            	    HttpStatus.OK.getReasonPhrase(),
            	    req.getRequestURI(),
            	    dWorkDayHoursSpecific
            	);

            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @GetMapping("/getDWorkDayHoursSpecificByDate/{date}")
    public Response getDWorkDayHoursSpecificByDate(final HttpServletRequest req, @PathVariable String date) throws ResourceNotFoundException {
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

                Optional<DWorkDayHoursSpesific> dWorkDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.getWorkDayHoursSpecificByDate(parsedDate);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    dWorkDayHoursSpecific
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (ParseException e) {
            throw new ResourceNotFoundException("Invalid date format, expected format is dd-MM-yyyy");
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @GetMapping("/getDWorkDayHoursSpecificListByDate/{date}")
    public Response getDWorkDayHoursSpecificListByDate(final HttpServletRequest req, @PathVariable String date) throws ResourceNotFoundException {
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

                List<DWorkDayHoursSpesific> dWorkDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.getWorkDayHoursSpecificListByDate(parsedDate);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    dWorkDayHoursSpecific
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (ParseException e) {
            throw new ResourceNotFoundException("Invalid date format, expected format is dd-MM-yyyy");
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @GetMapping("/getDWorkDayHoursSpecificByDateDesc/{date}/{description}")
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

                Optional<DWorkDayHoursSpesific> workDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.getWorkDayHoursSpecificByDateDesc(parsedDate, description);

              response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        workDayHoursSpecific
                    );
                
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        } 

        return response;
    }
    
    @GetMapping("/getWorkDayHoursByMonthYear/{month}/{year}")
    public Response getWorkDayHoursByMonthYear(final HttpServletRequest req, 
                                                @PathVariable int month, 
                                                @PathVariable int year) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        // Verifikasi header Authorization dan token JWT
        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        Response response;
        try {
            // Verifikasi token JWT dan ambil subject (username atau user ID)
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                .build()
                .verify(token)
                .getSubject();

            if (user != null) {
                // Panggil service untuk mengambil data berdasarkan bulan dan tahun
                List<DWorkDayHoursSpesific> workDayHoursSpecificList = 
                		dWorkDayHoursSpecificServiceImpl.getWorkDayHoursByMonthAndYear(month, year);

                // Membentuk response sukses
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        workDayHoursSpecificList
                    );
                
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/saveDWorkDayHoursSpecific")
    public Response saveDWorkDayHoursSpecific(final HttpServletRequest req, @RequestBody DWorkDayHoursSpesific dWorkDayHours) throws ResourceNotFoundException {
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
            	DWorkDayHoursSpesific workDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.saveWorkDayHoursSpecific(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDayHoursSpecific
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/updateDWorkDayHoursSpecific")
    public Response updateDWorkDayHoursSpecific(final HttpServletRequest req, @RequestBody DWorkDayHoursSpesific dWorkDayHours) throws ResourceNotFoundException {
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
            	DWorkDayHoursSpesific workDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.updateWorkDayHoursSpecific(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDayHoursSpecific
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/deleteDWorkDayHoursSpecific")
    public Response deleteDWorkDayHoursSpecific(final HttpServletRequest req, @RequestBody DWorkDayHoursSpesific dWorkDayHours) throws ResourceNotFoundException {
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
            	DWorkDayHoursSpesific workDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.deleteWorkDayHoursSpecific(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDayHoursSpecific
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @PostMapping("/restoreDWorkDayHoursSpecific")
    public Response restoreDWorkDayHoursSpecific(final HttpServletRequest req, @RequestBody DWorkDayHoursSpesific dWorkDayHours) throws ResourceNotFoundException {
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
            	DWorkDayHoursSpesific workDayHoursSpecific = dWorkDayHoursSpecificServiceImpl.restoreWorkDayHoursSpecific(dWorkDayHours); 

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDayHoursSpecific
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @PostMapping("/updateShiftTimes/{startTime}/{endTime}/{parsedDate}/{description}/{shift}")
    public Response updateShiftTimes(final HttpServletRequest req,
                                      @PathVariable("startTime") String startTime,
                                      @PathVariable("endTime") String endTime,
                                      @PathVariable("parsedDate") String parsedDate,
                                      @PathVariable("description") String description,
                                      @PathVariable("shift") int shift) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        // Check if Authorization header is present and valid
        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            // Extract the user from the JWT token
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                .build()
                .verify(token)
                .getSubject();

            if (user != null) {
                // Call service method to update shift times
                Optional<DWorkDayHoursSpesific> updatedWorkDayHours = dWorkDayHoursSpecificServiceImpl.updateShiftTimes(
                    startTime,
                    endTime,
                    parsedDate,
                    description,
                    shift
                );

                // Return response based on the result
                if (updatedWorkDayHours.isPresent()) {
                    DWorkDayHoursSpesific workDayHoursSpecific = updatedWorkDayHours.get();
                    response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        workDayHoursSpecific
                    );
                } else {
                    response = new Response(
                        new Date(),
                        HttpStatus.NOT_FOUND.value(),
                        null,
                        "No record found for the provided date",
                        req.getRequestURI(),
                        null
                    );
                }
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        
        
        return response;
    }
    
    @PostMapping("/saveDWorkDayHoursSpecificExcel")
    public Response saveDWorkDayHoursSpecificExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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

                dWorkDayHoursSpecificServiceImpl.deleteAllWorkDayHoursSpecific();
                List<DWorkDayHoursSpesific> dWorkDayHoursList = new ArrayList<>();

                try (InputStream inputStream = file.getInputStream()) {
                    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                    XSSFSheet sheet = workbook.getSheetAt(0);

                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);

                        if (row != null) {
                            DWorkDayHoursSpesific dWorkDayHours = new DWorkDayHoursSpesific();

                            // Parsing ID
                            Cell idCell = row.getCell(1);
                            if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                                dWorkDayHours.setDETAIL_WD_HOURS_SPECIFIC_ID(BigDecimal.valueOf(idCell.getNumericCellValue()));
                            } else {
                                continue; // Skip row if ID is missing or invalid
                            }

                            // Parsing tanggal
                            Cell dateCell = row.getCell(2);
                            if (dateCell != null) {
                                try {
                                    LocalDate workDate = parseDateCell(dateCell);
                                    if (workDate != null) {
                                        dWorkDayHours.setDATE_WD(Date.from(workDate.atStartOfDay(ZoneId.of("UTC")).toInstant()));
                                    } else {
                                        continue; // Skip if date is invalid
                                    }
                                } catch (Exception e) {
                                    System.out.println("Error parsing date for row " + (i + 1));
                                    continue; // Skip if there's an error parsing the date
                                }
                            } else {
                                System.out.println("Empty date cell at row " + (i + 1));
                                continue; // Skip if there's no date
                            }

                         // Parsing shift times and calculating total time
                            for (int j = 3; j <= 11; j++) {
                                Cell valueCell = row.getCell(j);
                                if (valueCell != null) {
                                	System.out.println("Cell " + (j + 1) + " value at row " + (i + 1) + ": " + valueCell.toString());
                                    try {
                                        switch (j) {
                                            case 3: // Start time for SHIFT 1
                                                dWorkDayHours.setSHIFT1_START_TIME(valueCell != null && !valueCell.getStringCellValue().trim().isEmpty() ? valueCell.getStringCellValue().trim() : null);
                                                break;

                                            case 4: // End time for SHIFT 1
                                                dWorkDayHours.setSHIFT1_END_TIME(valueCell != null && !valueCell.getStringCellValue().trim().isEmpty() ? valueCell.getStringCellValue().trim() : null);
                                                break;

                                            case 5: // Total time for SHIFT 1
                                            	dWorkDayHours.setSHIFT1_TOTAL_TIME(calculateTotalTime(
                                            	        dWorkDayHours.getSHIFT1_START_TIME(),
                                            	        dWorkDayHours.getSHIFT1_END_TIME()
                                            	    ));                                                break;

                                            case 6: // Start time for SHIFT 2
                                                dWorkDayHours.setSHIFT2_START_TIME(valueCell != null && !valueCell.getStringCellValue().trim().isEmpty() ? valueCell.getStringCellValue().trim() : null);
                                                break;

                                            case 7: // End time for SHIFT 2
                                                dWorkDayHours.setSHIFT2_END_TIME(valueCell != null && !valueCell.getStringCellValue().trim().isEmpty() ? valueCell.getStringCellValue().trim() : null);
                                                break;

                                            case 8: // Total time for SHIFT 2
                                            	dWorkDayHours.setSHIFT2_TOTAL_TIME(calculateTotalTime(
                                            	        dWorkDayHours.getSHIFT2_START_TIME(),
                                            	        dWorkDayHours.getSHIFT2_END_TIME()
                                            	    ));                                                break;

                                            case 9: // Start time for SHIFT 3
                                                dWorkDayHours.setSHIFT3_START_TIME(valueCell != null && !valueCell.getStringCellValue().trim().isEmpty() ? valueCell.getStringCellValue().trim() : null);
                                                break;

                                            case 10: // End time for SHIFT 3
                                                dWorkDayHours.setSHIFT3_END_TIME(valueCell != null && !valueCell.getStringCellValue().trim().isEmpty() ? valueCell.getStringCellValue().trim() : null);
                                                break;

                                            case 11: // Total time for SHIFT 3
                                            	dWorkDayHours.setSHIFT3_TOTAL_TIME(calculateTotalTime(
                                            	        dWorkDayHours.getSHIFT3_START_TIME(),
                                            	        dWorkDayHours.getSHIFT3_END_TIME()
                                            	    ));                                                break;

                                            default:
                                                break;
                                        }
                                        
                                    } catch (Exception e) {
                                        System.out.println("Error processing cell " + j + ": " + e.getMessage());
                                    }
                                }
                            }


                            // Parsing description
                            Cell descriptionCell = row.getCell(12);
                            if (descriptionCell != null && descriptionCell.getCellType() == CellType.STRING) {
                                dWorkDayHours.setDESCRIPTION(descriptionCell.getStringCellValue());
                            }

                            // Set default status and timestamps
                            dWorkDayHours.setSTATUS(BigDecimal.valueOf(1)); // Active status
                            dWorkDayHours.setCREATION_DATE(new Date());
                            dWorkDayHours.setLAST_UPDATE_DATE(new Date());

                            dWorkDayHoursSpecificServiceImpl.saveWorkDayHoursSpecific(dWorkDayHours);
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

    // Method to parse date cells
    private LocalDate parseDateCell(Cell dateCell) {
        if (dateCell.getCellType() == CellType.NUMERIC) {
            return dateCell.getLocalDateTimeCellValue().toLocalDate();
        } else if (dateCell.getCellType() == CellType.STRING) {
            try {
                return LocalDate.parse(dateCell.getStringCellValue(), DateTimeFormatter.ofPattern("dd/MM/yyyy"));
            } catch (DateTimeParseException e) {
                System.out.println("Error parsing date: " + e.getMessage());
                return null;
            }
        }
        return null; // Return null if the cell is not valid for date parsing
    }
    
    private BigDecimal calculateTotalTime(String startTimeStr, String endTimeStr) {
        // Jika salah satu parameter kosong atau null, kembalikan 0
        if (startTimeStr == null || startTimeStr.isEmpty() || endTimeStr == null || endTimeStr.isEmpty()) {
            return BigDecimal.ZERO;
        }

        try {
            int startMinutes = combineTimeWithMinutes(startTimeStr);
            int endMinutes = combineTimeWithMinutes(endTimeStr);
            
            int diffMinutes = endMinutes - startMinutes;
            
            if (diffMinutes < 0) {
                diffMinutes += 24 * 60; 
            }

            return BigDecimal.valueOf(diffMinutes);
        } catch (Exception e) {
            System.out.println("Error calculating total time: " + e.getMessage());
            return BigDecimal.ZERO; // Default to 0 if any exception occurs
        }
    }


    
    private int combineTimeWithMinutes(String timeStr) {
        String[] parts = timeStr.split(":");
        int hours = Integer.parseInt(parts[0]);
        int minutes = Integer.parseInt(parts[1]);
        return hours * 60 + minutes;
    }

    
    @RequestMapping("/exportDWorkDayHoursSpecificExcel")
    public ResponseEntity<InputStreamResource> exportDWorkDayHoursSpecificExcel() throws IOException {
        String filename = "MASTER_WORKDAY_HOURS.xlsx";

        ByteArrayInputStream data = dWorkDayHoursSpecificServiceImpl.exportWorkDayHoursSpecificExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
   
}
