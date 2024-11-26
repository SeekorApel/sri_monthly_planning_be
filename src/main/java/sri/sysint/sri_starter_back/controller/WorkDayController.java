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
import java.util.Locale;
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
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.WorkDay; // Import your WorkDay model
import sri.sysint.sri_starter_back.service.WorkDayServiceImpl; // Import your WorkDay service

@CrossOrigin(maxAge = 3600)
@RestController
public class WorkDayController {

    private Response response;

    @Autowired
    private WorkDayServiceImpl workDayServiceImpl;

    @PersistenceContext
    private EntityManager em;

    @GetMapping("/getAllWorkDays")
    public Response getAllWorkDays(final HttpServletRequest req) throws ResourceNotFoundException {
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
                List<WorkDay> workDays = workDayServiceImpl.getAllWorkDays();

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDays
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }
    
    @GetMapping("/getAllWorkDaysByDateRange/{startDate}/{endDate}")
    public Response getAllWorkDaysByDateRange(final HttpServletRequest req, 
                                              @PathVariable String startDate, 
                                              @PathVariable String endDate) throws ResourceNotFoundException {
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
                Date parsedStartDate = dateFormat.parse(startDate);
                Date parsedEndDate = dateFormat.parse(endDate);

                List<WorkDay> workDays = workDayServiceImpl.getAllWorkDaysByDateRange(parsedStartDate, parsedEndDate);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDays
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }


    @GetMapping("/getWorkDayByDate/{date}")
    public Response getWorkDayByDate(final HttpServletRequest req, @PathVariable String date) throws ResourceNotFoundException {
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
                
                Optional<WorkDay> workDay = workDayServiceImpl.getWorkDayByDate(parsedDate);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    workDay
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/saveWorkDay")
    public Response saveWorkDay(final HttpServletRequest req, @RequestBody WorkDay workDay) throws ResourceNotFoundException {
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
                WorkDay savedWorkDay = workDayServiceImpl.saveWorkDay(workDay);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    savedWorkDay
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/updateWorkDay")
    public Response updateWorkDay(final HttpServletRequest req, @RequestBody WorkDay workDay) throws ResourceNotFoundException {
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
                WorkDay updatedWorkDay = workDayServiceImpl.updateWorkDay(workDay);

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

    @PostMapping("/deleteWorkDay")
    public Response deleteWorkDay(final HttpServletRequest req, @RequestBody WorkDay workDay) throws ResourceNotFoundException {
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
                WorkDay deletedWorkDay = workDayServiceImpl.deleteWorkDay(workDay);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    deletedWorkDay
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/restoreWorkDay")
    public Response restoreWorkDay(final HttpServletRequest req, @RequestBody WorkDay workDay) throws ResourceNotFoundException {
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
                WorkDay restoredWorkDay = workDayServiceImpl.restoreWorkDay(workDay);

                response = new Response(
                    new Date(),
                    HttpStatus.OK.value(),
                    null,
                    HttpStatus.OK.getReasonPhrase(),
                    req.getRequestURI(),
                    restoredWorkDay
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

//    @PostMapping("/saveWorkDaysExcel")
//    public Response saveWorkDaysExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
//        String header = req.getHeader("Authorization");
//
//        if (header == null || !header.startsWith("Bearer ")) {
//            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
//        }
//
//        String token = header.replace("Bearer ", "");
//        Response response;
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
//                workDayServiceImpl.deleteAllWorkDays();  
//
//                try (InputStream inputStream = file.getInputStream()) {
//                    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
//                    XSSFSheet sheet = workbook.getSheetAt(0);
//
//                    List<WorkDay> workDays = new ArrayList<>();
//
//                    Row headerRow = sheet.getRow(0);
//                    if (headerRow == null) {
//                        return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Invalid file format", req.getRequestURI(), null);
//                    }
//
//                    // Date format for "Wed July 03, 2024"
//                    DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("E MMMM dd, yyyy", Locale.ENGLISH);
//                    // Date format for ZonedDateTime parsing
//                    DateTimeFormatter zonedDateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSSZ");
//
//                    for (int j = 1; j < headerRow.getLastCellNum(); j++) {
//                        Cell dateCell = headerRow.getCell(j);
//                        if (dateCell != null) {
//                            ZonedDateTime zonedDateTime = null;
//
//                            try {
//                                if (dateCell.getCellType() == CellType.NUMERIC) {
//                                    if (DateUtil.isCellDateFormatted(dateCell)) {
//                                        LocalDate localDate = dateCell.getDateCellValue().toInstant()
//                                                .atZone(ZoneId.systemDefault()) 
//                                                .toLocalDate();
//                                        zonedDateTime = localDate.atStartOfDay(ZoneId.of("UTC"));
//                                    } else {
//                                        return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Invalid date format in file: " + dateCell.toString(), req.getRequestURI(), null);
//                                    }
//                                } else if (dateCell.getCellType() == CellType.STRING) {
//                                    try {
//                                        // Attempt to parse the string date
//                                        LocalDate parsedDate = LocalDate.parse(dateCell.getStringCellValue().trim(), dateFormatter);
//                                        zonedDateTime = parsedDate.atStartOfDay(ZoneId.of("UTC"));
//                                    } catch (Exception e) {
//                                        return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Error parsing date: " + dateCell.toString(), req.getRequestURI(), null);
//                                    }
//                                }
//                            } catch (Exception e) {
//                                return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Error processing date cell: " + dateCell.toString(), req.getRequestURI(), null);
//                            }
//
//                            Date date = Date.from(zonedDateTime.toInstant());
//                            WorkDay workDay = new WorkDay();
//                            workDay.setDATE_WD(date);
//
//                            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//                                Row row = sheet.getRow(i);
//                                if (row == null || isRowEmpty(row)) {
//                                    break; 
//                                }
//
//                                Cell cell = row.getCell(j);  
//                                String shiftName = row.getCell(0).getStringCellValue();  
//                                int status = 0;
//
//                                if (cell != null) {
//                                    System.out.println("Row: " + (i + 1) + ", Col: " + (j + 1) + ", Cell Type: " + cell.getCellType() + ", Value: " + cell.toString());
//
//                                    if (cell.getCellType() == CellType.STRING) {
//                                        String cellValue = cell.getStringCellValue().trim();
//                                        if ("v".equalsIgnoreCase(cellValue)) {
//                                            status = 1;
//                                        } else if (cellValue.isEmpty()) {
//                                            status = 0;
//                                        }
//                                    } else if (cell.getCellType() == CellType.BLANK) {
//                                        status = 0;
//                                    }
//                                } else {
//                                    status = 0; 
//                                }
//
//                                switch (shiftName.toUpperCase()) {
//                                    case "WD_SHIFT_1":
//                                        workDay.setIWD_SHIFT_1(BigDecimal.valueOf(status));
//                                        break;
//                                    case "WD_SHIFT_2":
//                                        workDay.setIWD_SHIFT_2(BigDecimal.valueOf(status));
//                                        break;
//                                    case "WD_SHIFT_3":
//                                        workDay.setIWD_SHIFT_3(BigDecimal.valueOf(status));
//                                        break;
//                                    case "OT_TL_1":
//                                        workDay.setIOT_TL_1(BigDecimal.valueOf(status));
//                                        break;
//                                    case "OT_TL_2":
//                                        workDay.setIOT_TL_2(BigDecimal.valueOf(status));
//                                        break;
//                                    case "OT_TL_3":
//                                        workDay.setIOT_TL_3(BigDecimal.valueOf(status));
//                                        break;
//                                    case "OT_TT_1":
//                                        workDay.setIOT_TT_1(BigDecimal.valueOf(status));
//                                        break;
//                                    case "OT_TT_2":
//                                        workDay.setIOT_TT_2(BigDecimal.valueOf(status));
//                                        break;
//                                    case "OT_TT_3":
//                                        workDay.setIOT_TT_3(BigDecimal.valueOf(status));
//                                        break;
//                                    default:
//                                        throw new IllegalArgumentException("Unknown shift name: " + shiftName);
//                                }
//                            }
//
//                            int iwdShift1Status = workDay.getIWD_SHIFT_1().intValue();
//                            int iwdShift2Status = workDay.getIWD_SHIFT_2().intValue();
//                            int iwdShift3Status = workDay.getIWD_SHIFT_3().intValue();
//
//                            if (iwdShift1Status == 0 && iwdShift2Status == 0 && iwdShift3Status == 0) {
//                                workDay.setOFF(BigDecimal.ONE); 
//                                workDay.setSEMI_OFF(BigDecimal.ZERO);
//                            } else if (iwdShift1Status == 1 && iwdShift2Status == 1 && iwdShift3Status == 1) {
//                                workDay.setOFF(BigDecimal.ZERO); 
//                                workDay.setSEMI_OFF(BigDecimal.ZERO);
//                            } else {
//                                workDay.setOFF(BigDecimal.ZERO); 
//                                workDay.setSEMI_OFF(BigDecimal.ONE);
//                            }
//
//                            workDayServiceImpl.saveWorkDay(workDay); 
//                            workDays.add(workDay);
//                        }
//                    }
//
//                    response = new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), workDays);
//
//                } catch (IOException e) {
//                    response = new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file", req.getRequestURI(), null);
//                }
//            } else {
//                throw new ResourceNotFoundException("User not found");
//            }
//        } catch (Exception e) {
//            throw new ResourceNotFoundException("JWT token is not valid or expired");
//        }
//
//        return response;
//    }
    
    @PostMapping("/saveWorkDaysExcel")
    public Response saveWorkDaysExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {

        if (file.isEmpty()) {
            return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "No file uploaded", req.getRequestURI(), null);
        }

        workDayServiceImpl.deleteAllWorkDays();

        try (InputStream inputStream = file.getInputStream()) {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);

            List<WorkDay> workDays = new ArrayList<>();
            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("EEEE, dd MMMM yyyy", Locale.ENGLISH);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isRowEmpty(row)) {
                    continue;
                }

                Cell dateCell = row.getCell(0);
                if (dateCell == null || (dateCell.getCellType() != CellType.STRING && dateCell.getCellType() != CellType.NUMERIC)) {
                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Invalid date format in file", req.getRequestURI(), null);
                }

                ZonedDateTime zonedDateTime = null;

                try {
                    if (dateCell.getCellType() == CellType.NUMERIC) {
                        if (DateUtil.isCellDateFormatted(dateCell)) {
                            LocalDate localDate = dateCell.getDateCellValue().toInstant()
                                    .atZone(ZoneId.systemDefault()) 
                                    .toLocalDate();
                            zonedDateTime = localDate.atStartOfDay(ZoneId.of("UTC"));
                        } else {
                            return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Invalid date format in file: " + dateCell.toString(), req.getRequestURI(), null);
                        }
                    } else if (dateCell.getCellType() == CellType.STRING) {
                        LocalDate parsedDate = LocalDate.parse(dateCell.getStringCellValue().trim(), dateFormatter);
                        zonedDateTime = parsedDate.atStartOfDay(ZoneId.of("UTC"));
                    }
                } catch (Exception e) {
                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, "Error parsing date: " + dateCell.toString(), req.getRequestURI(), null);
                }

                Date date = Date.from(zonedDateTime.toInstant());

                WorkDay workDay = new WorkDay();
                workDay.setDATE_WD(date);

                for (int j = 1; j <= 10; j++) {
                    Cell cell = row.getCell(j);
                    String shiftName = sheet.getRow(0).getCell(j).getStringCellValue();
                    int status = 0;

                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue().trim();
                            if ("v".equalsIgnoreCase(cellValue)) {
                                status = 1;
                            }
                        } else if (cell.getCellType() == CellType.BLANK) {
                            status = 0;
                        }
                    }

                    switch (shiftName.toUpperCase()) {
                        case "WD_SHIFT_1":
                            workDay.setIWD_SHIFT_1(BigDecimal.valueOf(status));
                            break;
                        case "WD_SHIFT_2":
                            workDay.setIWD_SHIFT_2(BigDecimal.valueOf(status));
                            break;
                        case "WD_SHIFT_3":
                            workDay.setIWD_SHIFT_3(BigDecimal.valueOf(status));
                            break;
                        case "OT_TL_1":
                            workDay.setIOT_TL_1(BigDecimal.valueOf(status));
                            break;
                        case "OT_TL_2":
                            workDay.setIOT_TL_2(BigDecimal.valueOf(status));
                            break;
                        case "OT_TL_3":
                            workDay.setIOT_TL_3(BigDecimal.valueOf(status));
                            break;
                        case "OT_TT_1":
                            workDay.setIOT_TT_1(BigDecimal.valueOf(status));
                            break;
                        case "OT_TT_2":
                            workDay.setIOT_TT_2(BigDecimal.valueOf(status));
                            break;
                        case "OT_TT_3":
                            workDay.setIOT_TT_3(BigDecimal.valueOf(status));
                            break;
                        case "OFF":
                            workDay.setOFF(BigDecimal.valueOf(status));
                            break;
                        default:
                            throw new IllegalArgumentException("Unknown shift name: " + shiftName);
                    }
                    workDay.setSEMI_OFF(BigDecimal.ZERO);
                }

                workDayServiceImpl.saveWorkDay(workDay);
                workDays.add(workDay);
            }

            return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), workDays);

        } catch (IOException e) {
            return new Response(new Date(), HttpStatus.INTERNAL_SERVER_ERROR.value(), null, "Error processing file", req.getRequestURI(), null);
        }
    }

  
    private boolean isRowEmpty(Row row) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false; 
            }
        }
        return true; 
    }

    @RequestMapping("/exportWorkDaysExcel")
    public ResponseEntity<InputStreamResource> exportWorkDaysExcel() throws IOException {
        String filename = "MASTER_WORK_DAY.xlsx";

        ByteArrayInputStream data = workDayServiceImpl.exportWDsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
    @PostMapping("/turnOnOvertime/{dateWd}")
    public Response turnOnOvertime(final HttpServletRequest req, @PathVariable String dateWd) throws ResourceNotFoundException {
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
                WorkDay updatedWorkDay = workDayServiceImpl.turnOnOvertime(dateWd);

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
    
    @PostMapping("/turnOnShift/{dateWd}/{shift}")
    public Response turnOnShift(final HttpServletRequest req, @PathVariable String dateWd, @PathVariable String shift) throws ResourceNotFoundException {
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
                WorkDay updatedWorkDay = workDayServiceImpl.turnOnShift(dateWd, shift);

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

    @PostMapping("/turnOffShift/{dateWd}/{shift}")
    public Response turnOffShift(final HttpServletRequest req, @PathVariable String dateWd, @PathVariable String shift) throws ResourceNotFoundException {
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
                WorkDay updatedWorkDay = workDayServiceImpl.turnOffShift(dateWd, shift);

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
