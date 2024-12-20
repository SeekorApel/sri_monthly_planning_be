package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

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
import sri.sysint.sri_starter_back.model.CTAssy;
import sri.sysint.sri_starter_back.model.CTCuring;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.MachineCuring;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.CTCuringServiceImpl;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.MachineCuringRepo;

@CrossOrigin(maxAge = 3600)
@RestController
public class CTCuringController {
	
	private Response response;	

	@Autowired
	private CTCuringServiceImpl ctCuringServiceImpl;
	
    @Autowired
    private ItemCuringRepo itemCuringRepo;

    @Autowired
    private MachineCuringRepo machineCuringRepo;

	@PersistenceContext	
	private EntityManager em;
	
	@GetMapping("/getAllCTCuring")
	public Response getAllCTCuring(final HttpServletRequest req) throws ResourceNotFoundException {
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
	        	List<CTCuring> ctCurings = new ArrayList<>();
	        	ctCurings = ctCuringServiceImpl.getAllCTCuring();
		
			    response = new Response(
			        new Date(),
			        HttpStatus.OK.value(),
			        null,
			        HttpStatus.OK.getReasonPhrase(),
			        req.getRequestURI(),
			        ctCurings
			    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@GetMapping("/getCTCuringById/{id}")
	public Response getCTCuringById(final HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
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
	        	Optional<CTCuring> ctCuring = Optional.of(new CTCuring());
	        	ctCuring = ctCuringServiceImpl.getCTCuringById(id);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        ctCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveCTCuring")
	public Response saveCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring savedCTCuring = ctCuringServiceImpl.saveCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        savedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/updateCTCuring")
	public Response updateCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring updatedCTCuring = ctCuringServiceImpl.updateCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        updatedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/deleteCTCuring")
	public Response deleteCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring deletedCTCuring = ctCuringServiceImpl.deleteCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        deletedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/restoreCTCuring")
	public Response activateCTCuring(final HttpServletRequest req, @RequestBody CTCuring ctCuring) throws ResourceNotFoundException {
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
	        	CTCuring activatedCTCuring = ctCuringServiceImpl.activateCTCuring(ctCuring);

	    	    response = new Response(
	    	        new Date(),
	    	        HttpStatus.OK.value(),
	    	        null,
	    	        HttpStatus.OK.getReasonPhrase(),
	    	        req.getRequestURI(),
	    	        activatedCTCuring
	    	    );
	        } else {
	            throw new ResourceNotFoundException("User not found");
	        }
	    } catch (Exception e) {
	        throw new ResourceNotFoundException("JWT token is not valid or expired");
	    }

	    return response;
	}
	
	@PostMapping("/saveCTCuringExcel")
	public Response saveCTCuringExcelFile(@RequestParam("file") MultipartFile file, final HttpServletRequest req) throws ResourceNotFoundException {
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
					
					List<CTCuring> ctCurings = new ArrayList<>();
					List<String> errorMessages = new ArrayList<>();
					
					for (int i = 3; i <= sheet.getLastRowNum(); i++) {
						Row row = sheet.getRow(i);
						if (row != null) {
							CTCuring ctCuring = new CTCuring();

							boolean hasError = false;

							String[] tableData = {
								"WIP", "Group Counter", "Var Group Counter", "Sequence", "WCT", "Operation Short Text", 
								"Operation Unit", "Base Quantity", "Standard Value Unit", "CT SEKON/1PC", "CT HR/1000PC",
								"normal shift-1", "normal shift-2", "normal shift-3", "FRIDAY 4", "BTOL TOTAL NORMAL","TOTAL FRIDAY",
								"ALLOWANE normal shift-1", "ALLOWANE normal shift-2", "ALLOWANE normal shift-3", "ALLOWANE Total",
								"OPERATIONAL TIME normal shift-1", "OPERATIONAL TIME normal shift-2", "OPERATIONAL TIME normal shift-3", "OPERATIONAL TIME Jumat shift-1", "OPERATIONAL TIME TOTAL NORMAL","OPERATIONAL TIME TOTAL FRIDAY",
								"KAPASITAS NORMAL SHIFT 1", "KAPASITAS NORMAL SHIFT 2", "KAPASITAS NORMAL SHIFT 3", "KAPASITAS JUMAT Shift 1", "KAPASITAS TOTAL NORMAL", 
								"KAPASITAS TOTAL FRIDAY","KAPASITAS	WAKTU TOTAL/CT","KAPASITAS WAKTU TOTAL/CT FRIDAY"
							};


							// Loop through columns to validate cells
							for (int col = 0; col <= 34; col++) {
								Cell cell = row.getCell(col);
								if (cell == null || cell.getCellType() == CellType.BLANK) {
									errorMessages.add("Data Tidak Valid, Terdapat Data Kosong pada Baris " + (i + 1) + " Kolom " + (col + 1)+" ( " + tableData[col] + " )");
									hasError = true;
								}
							}

							// Skip processing the row if there are errors
							if (hasError) {
								continue;
							}

							// Proceed with processing the row
							Cell wipCell = row.getCell(0);
							Cell operationShortTextCell = row.getCell(5);

							String wip = wipCell.getStringCellValue();
							String operationShortText = operationShortTextCell.getStringCellValue();
							Optional<ItemCuring> wipOpt = itemCuringRepo.findById(wip);
							Optional<MachineCuring> operationShortTextOpt = machineCuringRepo.findById(operationShortText);

							if (wipOpt.isPresent() && operationShortTextOpt.isPresent()) {

								ctCuring.setCT_CURING_ID(ctCuringServiceImpl.getNewId());
								ctCuring.setWIP(wip);

								for (int col = 1; col <= 34; col++) {
									Cell cell = row.getCell(col);
									switch (col) {
										case 1:
											ctCuring.setGROUP_COUNTER(getStringFromCell(cell));
											break;
										case 2:
											ctCuring.setVAR_GROUP_COUNTER(getStringFromCell(cell));
											break;
										case 3:
											ctCuring.setSEQUENCE(getBigDecimalFromCell(cell));
											break;
										case 4:
											ctCuring.setWCT(getStringFromCell(cell));
											break;
										case 5:
											ctCuring.setOPERATION_SHORT_TEXT(getStringFromCell(cell));
											break;
										case 6:
											ctCuring.setOPERATION_UNIT(getStringFromCell(cell));
											break;
										case 7:
											ctCuring.setBASE_QUANTITY(getBigDecimalFromCell(cell));
											break;
										case 8:
											ctCuring.setSTANDART_VALUE_UNIT(getStringFromCell(cell));
											break;
										case 9:
											ctCuring.setCT_SEC1(getBigDecimalFromCell(cell));
											break;
										case 10:
											ctCuring.setCT_HR1000(getBigDecimalFromCell(cell));
											break;
										case 11:
											ctCuring.setWH_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
											break;
										case 12:
											ctCuring.setWH_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
											break;
										case 13:
											ctCuring.setWH_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
											break;
										case 14:
											ctCuring.setWH_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
											break;
										case 15:
											ctCuring.setWH_TOTAL_NORMAL_SHIFT(getBigDecimalFromCell(cell));
											break;
										case 16:
											ctCuring.setWH_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
											break;
										case 17:
											ctCuring.setALLOW_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
											break;
										case 18:
											ctCuring.setALLOW_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
											break;
										case 19:
											ctCuring.setALLOW_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
											break;
										case 20:
											ctCuring.setALLOW_TOTAL(getBigDecimalFromCell(cell));
											break;
										case 21:
											ctCuring.setOP_TIME_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
											break;
										case 22:
											ctCuring.setOP_TIME_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
											break;
										case 23:
											ctCuring.setOP_TIME_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
											break;
										case 24:
											ctCuring.setOP_TIME_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
											break;
										case 25:
											ctCuring.setOP_TIME_NORMAL_SHIFT(getBigDecimalFromCell(cell));
											break;
										case 26:
											ctCuring.setOP_TIME_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
											break;
										case 27:
											ctCuring.setKAPS_NORMAL_SHIFT_0(getBigDecimalFromCell(cell));
											break;
										case 28:
											ctCuring.setKAPS_NORMAL_SHIFT_1(getBigDecimalFromCell(cell));
											break;
										case 29:
											ctCuring.setKAPS_NORMAL_SHIFT_2(getBigDecimalFromCell(cell));
											break;
										case 30:
											ctCuring.setKAPS_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
											break;
										case 31:
											ctCuring.setKAPS_TOTAL_NORMAL_SHIFT(getBigDecimalFromCell(cell));
											break;
										case 32:
											ctCuring.setKAPS_TOTAL_SHIFT_FRIDAY(getBigDecimalFromCell(cell));
											break;
										case 33:
											ctCuring.setWAKTU_TOTAL_CT_NORMAL(getBigDecimalFromCell(cell));
											break;
										case 34:
											ctCuring.setWAKTU_TOTAL_CT_FRIDAY(getBigDecimalFromCell(cell));
											break;
									}
								}

								ctCuring.setSTATUS(BigDecimal.valueOf(1));
								ctCuring.setCREATION_DATE(new Date());
								ctCuring.setLAST_UPDATE_DATE(new Date());
								ctCurings.add(ctCuring);
							} else {
								errorMessages.add("Data Tidak Valid, Data WIP pada Baris " + (i + 1) + " Tidak Ditemukan");
							}
						}
					}
					if (!errorMessages.isEmpty()) {
	                    return new Response(new Date(), HttpStatus.BAD_REQUEST.value(), null, String.join("; ", errorMessages), req.getRequestURI(), null);
	                }
					ctCuringServiceImpl.deleteAllCTCuring();
					for(CTCuring ctCuring : ctCurings){
						ctCuringServiceImpl.saveCTCuring(ctCuring);
					}

					return new Response(new Date(), HttpStatus.OK.value(), null, "File processed and data saved", req.getRequestURI(), ctCurings);

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

	private boolean isRowEmpty(Row row) {
	    for (int j = 0; j < row.getLastCellNum(); j++) {
	        Cell cell = row.getCell(j);
	        if (cell != null && cell.getCellType() != CellType.BLANK) {
	            return false;
	        }
	    }
	    return true;
	}

	private BigDecimal getBigDecimalFromCell(Cell cell) {
	    if (cell == null) {
	        return null;
	    }
	    switch (cell.getCellType()) {
	        case STRING:
	            String value = cell.getStringCellValue();
	            try {
	                return new BigDecimal(value);
	            } catch (NumberFormatException e) {
	                System.out.println("Invalid number format in cell: " + value);
	                return null; 
	            }
	        case NUMERIC:
	            return BigDecimal.valueOf(cell.getNumericCellValue());
	        default:
	            return null; 
	    }
	}

	private String getStringFromCell(Cell cell) {
	    if (cell == null) {
	        return null;
	    }
	    return cell.getStringCellValue();
	}
	
    @GetMapping("/exportCTCuringExcel")
    public ResponseEntity<InputStreamResource> exportCTCuringExcel() throws IOException {
        String filename = "MASTER_CT_CURING_DATA.xlsx";

        // Generate the Excel data using the service method
        ByteArrayInputStream data = ctCuringServiceImpl.exportCTCuringsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
    
	@GetMapping("/layoutCTCuringsExcel")
    public ResponseEntity<InputStreamResource> layoutCTCuringsExcel() throws IOException {
        String filename = "LAYOUT_MASTER_CT_CURING.xlsx";
        ByteArrayInputStream data = ctCuringServiceImpl.layoutCTCuringsExcel();
        InputStreamResource file = new InputStreamResource(data);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
}
