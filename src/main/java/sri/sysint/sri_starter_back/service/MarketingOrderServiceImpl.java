package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.repository.MarketingOrderRepo;
import sri.sysint.sri_starter_back.repository.HeaderMarketingOrderRepo;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.repository.ProductTypeRepo;
import sri.sysint.sri_starter_back.repository.SettingRepo;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.MarketingOrder;
import sri.sysint.sri_starter_back.model.DetailMarketingOrder;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.ProductType;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.model.transaksi.EditMarketingOrderMarketing;
import sri.sysint.sri_starter_back.model.transaksi.SaveMarketingOrderPPC;
import sri.sysint.sri_starter_back.model.view.ViewDetailMarketingOrder;
import sri.sysint.sri_starter_back.model.view.ViewHeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.view.ViewMarketingOrder;
import sri.sysint.sri_starter_back.repository.DetailMarketingOrderRepo;


@Service
@Transactional
public class MarketingOrderServiceImpl {
	
	@Autowired
    private MarketingOrderRepo marketingOrderRepo;
	
	@Autowired
    private HeaderMarketingOrderRepo headerMarketingOrderRepo;
	
	@Autowired
    private DetailMarketingOrderRepo detailMarketingOrderRepo;
	
	@Autowired
    private ItemCuringRepo itemCuringRepo;
	
	@Autowired
    private ProductRepo productRepo;
	
	@Autowired
    private SettingRepo settingRepo;
	
	@Autowired
    private ProductTypeRepo productTypeRepo;
	
	public MarketingOrderServiceImpl(MarketingOrderRepo marketingOrderRepo, HeaderMarketingOrderRepo headerMarketingOrderRepo, DetailMarketingOrderRepo detailMarketingOrderRepo){
        this.marketingOrderRepo = marketingOrderRepo;
        this.headerMarketingOrderRepo = headerMarketingOrderRepo;
        this.detailMarketingOrderRepo = detailMarketingOrderRepo;
    }
	
	
	//GET NEW ID HEADER MO
    public BigDecimal getNewHeaderMarketingOrderId() {
        return headerMarketingOrderRepo.getNewId().add(BigDecimal.valueOf(1));
    }

	//GET NEW ID DETAIL MO
    public BigDecimal getNewDetailMarketingOrderId() {
        return detailMarketingOrderRepo.getNewId().add(BigDecimal.valueOf(1));
    }
	
	//Add dicky
	//GET NEW ID MO
	public String getLastIdMo() {
	    String lastId = marketingOrderRepo.getLastIdMo();
	    if (lastId == null || lastId.isEmpty()) {
	        return "MO-001";
	    } else {
	        String prefix = lastId.substring(0, 3);
	        int number = Integer.parseInt(lastId.substring(3));
	        number++;
	        return String.format("%s%03d", prefix, number);
	    }
	}
	
	//GET CAPACITY FOR MO
    public String getCapacityValue() {
        return settingRepo.getCapacity();
    }

	//GET ALL MARKETING ORDER
    public List<MarketingOrder> getAllMarketingOrderLatest() {
        Iterable<MarketingOrder> moes = marketingOrderRepo.findLatestMarketingOrders();
        List<MarketingOrder> moList = new ArrayList<>();
        for (MarketingOrder item : moes) {
            MarketingOrder marketingOrderTemp = new MarketingOrder(item);
            moList.add(marketingOrderTemp);
        }
        return moList;
    }
    
    //GET ALL MARKETING ORDER LATEST BY ROLE
    public List<MarketingOrder> getAllMarketingOrderMarketing(String role) {
        if ("Marketing FED".equals(role)) {
            return marketingOrderRepo.findLatestMarketingOrderFED();
        } else if ("Marketing FDR".equals(role)) {
            return marketingOrderRepo.findLatestMarketingOrderFDR();
        }
        return new ArrayList<>(); 
    }

    
    //GET ALL MARKETING ORDER
    public List<MarketingOrder> getAllMarketingOrder(String month0, String month1, String month2, String type) {
        return marketingOrderRepo.findtMarketingOrders(month0, month1, month2, type);
    }
    
    //CHECK MONTH AVAILABLE
    public int checkMonthsAvailability(String month1, String month2, String month3, String year1, String year2, String year3, String type) {
        List<MarketingOrder> existingOrders = marketingOrderRepo.checktMarketingOrders(month1, month2, month3, year1, year2, year3, type);
        
        return existingOrders.isEmpty() ? 0 : 1; 
    }

	//GET MO BY ID
    public Optional<MarketingOrder> getMarketingOrderById(String id) {
        Optional<MarketingOrder> marketingOrder = marketingOrderRepo.findById(id);
        return marketingOrder;
    }
    
    //SAVE MARKETING ORDER, HEADER, DETAIL ROLE PPC
	public int saveMarketingOrderPPC(SaveMarketingOrderPPC marketingOrder) {
			
			int statusSave = 0;
			int statusMo = 0;
			int statusHmo = 0;
			int statusDmo = 0;
			
			SaveMarketingOrderPPC mo = new SaveMarketingOrderPPC(marketingOrder);
			
			//Save to SRI_IMPP_T_MARKETINGORDER
			try {
				MarketingOrder saveMo = new MarketingOrder(mo.getMarketingOrder());
				if (saveMo.getRevisionPpc() == null) {
				    saveMo.setRevisionPpc(BigDecimal.ZERO); // Set to 0 if null or zero
					saveMo.setStatusFilled(BigDecimal.ONE);

				} else {
				    saveMo.setRevisionPpc(saveMo.getRevisionPpc().add(BigDecimal.ONE));
				    saveMo.setStatusFilled(BigDecimal.valueOf(3));
				}

				saveMo.setStatus(BigDecimal.valueOf(1));  
				saveMo.setCreationDate(new Date());
				saveMo.setLastUpdateDate(new Date());
				MarketingOrder saveDb = marketingOrderRepo.save(saveMo);
				if(saveDb != null) {
					statusMo = 1;
				}
			}catch (Exception e){
	            System.err.println("Error saving MarketingOrder: " + e.getMessage());
	            throw e;
			}
			
			//Save to SRI_IMPP_D_HEADERMARKETINGORDER
			try {
		        List<HeaderMarketingOrder> headerMoList = mo.getHeaderMarketingOrder();
		        for (HeaderMarketingOrder headerMO : headerMoList) {
		        	HeaderMarketingOrder savedHeaderMO = new HeaderMarketingOrder(headerMO);
		        	savedHeaderMO.setHeaderId(getNewHeaderMarketingOrderId());
		        	savedHeaderMO.setStatus(BigDecimal.valueOf(1));
		        	savedHeaderMO.setCreationDate(new Date());
		        	savedHeaderMO.setLastUpdateDate(new Date());
		            headerMarketingOrderRepo.save(savedHeaderMO);
		            statusHmo = 1;
		        }
			}catch(Exception e) {
	            System.err.println("Error saving HeaderMO: " + e.getMessage());
	            throw e;
			}
			
			//Save to SRI_IMPP_D_MARKETINGORDER
			try {
				List<DetailMarketingOrder> detailMoList = mo.getDetailMarketingOrder();
		        for (DetailMarketingOrder detailMo : detailMoList) {
		        	DetailMarketingOrder savedDetailMo = new DetailMarketingOrder(detailMo);
		        	savedDetailMo.setDetailId(getNewDetailMarketingOrderId());
		        	savedDetailMo.setStatus(BigDecimal.valueOf(1)); 
		        	savedDetailMo.setCreationDate(new Date());
		        	savedDetailMo.setLastUpdateDate(new Date());
		            detailMarketingOrderRepo.save(savedDetailMo);
		        }
		        statusDmo = 1;
			}catch(Exception e) {
	            System.err.println("Error saving DetailMO: " + e.getMessage());
	            throw e;
			}
			
			if(statusMo == 1 && statusHmo == 1 && statusDmo == 1) {
				statusSave = 1;
			}
			
			statusSave = 1;
			
			return statusSave;
		}
	
	//End add dicky
    
    //GET HEADER MO BY ID
    public Optional<HeaderMarketingOrder> getHeaderMOById(BigDecimal id) {
        return headerMarketingOrderRepo.findById(id);
    }

    //GET/GENERATE DETAIL MARKETING ORDER FROM PARAM
    public List<ViewDetailMarketingOrder> getDetailMarketingOrders(
    	    BigDecimal totalHKTT1, BigDecimal totalHKTT2, BigDecimal totalHKTT3,
    	    BigDecimal totalHKTL1, BigDecimal totalHKTL2, BigDecimal totalHKTL3,
    	    String productMerk, String monthYear0, String monthYear1, String monthYear2) {
    	
	    	DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM-yyyy");
	
	    	LocalDate pre_monthYear0 = LocalDate.parse("01-" + monthYear0, DateTimeFormatter.ofPattern("dd-MM-yyyy"));
	    	LocalDate previousMonth0 = pre_monthYear0.minusMonths(1);  
	    	monthYear0 = previousMonth0.format(formatter);
	
	    	LocalDate pre_monthYear1 = LocalDate.parse("01-" + monthYear1, DateTimeFormatter.ofPattern("dd-MM-yyyy"));
	    	LocalDate previousMonth1 = pre_monthYear1.minusMonths(1);
	    	monthYear1 = previousMonth1.format(formatter);
	
	    	LocalDate pre_monthYear2 = LocalDate.parse("01-" + monthYear2, DateTimeFormatter.ofPattern("dd-MM-yyyy"));
	    	LocalDate previousMonth2 = pre_monthYear2.minusMonths(1);
	    	monthYear2 = previousMonth2.format(formatter);
	    	
	    	System.out.println("1: " +monthYear0);
	    	System.out.println("2: " +monthYear1);
	    	System.out.println("3: " +monthYear2);

	    	List<MarketingOrder> marketingOrder = marketingOrderRepo.findByMonth(monthYear0, monthYear1, monthYear2);

	    	List<DetailMarketingOrder> detailMarketingOrder = new ArrayList<>();
	    	if (marketingOrder != null && !marketingOrder.isEmpty()) {
	    	    for (MarketingOrder order : marketingOrder) {
	    	        detailMarketingOrder.addAll(detailMarketingOrderRepo.findByMoId(order.getMoId()));
	    	    }
	    	}

        	            
            List<Map<String, Object>> listData = detailMarketingOrderRepo.getDataTable(
	    	        totalHKTT1, totalHKTT2, totalHKTT3,
	    	        totalHKTL1, totalHKTL2, totalHKTL3,
	    	        productMerk
	    	    );

    	    List<ViewDetailMarketingOrder> productList = new ArrayList<>();
    	    
    	    for (Map<String, Object> rowData : listData) {
    	        ViewDetailMarketingOrder product = new ViewDetailMarketingOrder();
    	            	        
    	        product.setPartNumber(rowData.get("PART_NUMBER") != null ? new BigDecimal(rowData.get("PART_NUMBER").toString()) : BigDecimal.ZERO);
    	        product.setSpareMould(rowData.get("SPARE_MOULD") != null ? new BigDecimal(rowData.get("SPARE_MOULD").toString()) : BigDecimal.ZERO);
    	        product.setMouldMonthlyPlan(rowData.get("MOULD_MONTHLY_PLAN") != null ? new BigDecimal(rowData.get("MOULD_MONTHLY_PLAN").toString()) : BigDecimal.ZERO);
    	        product.setQtyPerMould(rowData.get("NUMBER_OF_MOULD") != null ? new BigDecimal(rowData.get("NUMBER_OF_MOULD").toString()) : BigDecimal.ZERO);
    	        product.setQtyPerRak(rowData.get("QTY_PER_RAK") != null ? new BigDecimal(rowData.get("QTY_PER_RAK").toString()) : BigDecimal.ZERO);
    	        product.setMaxCapMonth0(rowData.get("KAPASITAS_MAKSIMUM_1") != null ? new BigDecimal(rowData.get("KAPASITAS_MAKSIMUM_1").toString()) : BigDecimal.ZERO);
    	        product.setMaxCapMonth1(rowData.get("KAPASITAS_MAKSIMUM_2") != null ? new BigDecimal(rowData.get("KAPASITAS_MAKSIMUM_2").toString()) : BigDecimal.ZERO);
    	        product.setMaxCapMonth2(rowData.get("KAPASITAS_MAKSIMUM_3") != null ? new BigDecimal(rowData.get("KAPASITAS_MAKSIMUM_3").toString()) : BigDecimal.ZERO);
    	        product.setCategory(rowData.get("CATEGORY") != null ? rowData.get("CATEGORY").toString() : null);
    	        product.setDescription(rowData.get("DESCRIPTION") != null ? rowData.get("DESCRIPTION").toString() : null);
    	        product.setCapacity(rowData.get("KAPA_PER_MOULD") != null ? new BigDecimal(rowData.get("KAPA_PER_MOULD").toString()) : BigDecimal.ZERO);
    	        
    	        if(marketingOrder == null) {
    	        	product.setMinOrder(null);
    	        	product.setMachineType(null);
            	}else {
            		for(DetailMarketingOrder detail:detailMarketingOrder) {
            			if(rowData.get("PART_NUMBER").equals(detail.getPartNumber())) {
            				product.setMinOrder(detail.getMinOrder());
            	        	product.setMachineType(detail.getMachineType());
            			}
            		}
            	}
    	        productList.add(product);
    	    }

    	    List<ViewDetailMarketingOrder> detailMarketingOrders = productList; 

    	    return detailMarketingOrders; 
    	}
    
    //SAVE DETIL MO
    public DetailMarketingOrder saveDetailMO(DetailMarketingOrder detailMO) {
        try {
            detailMO.setDetailId(getNewDetailMarketingOrderId());
            detailMO.setStatus(BigDecimal.valueOf(1)); 
            detailMO.setCreationDate(new Date());
            detailMO.setLastUpdateDate(new Date());
            return detailMarketingOrderRepo.save(detailMO);
        } catch (Exception e) {
            System.err.println("Error saving DetailMO: " + e.getMessage());
            throw e;
        }
    }
  
    //GET ALL MARKETING ORDER BY MO ID (MO, HEADER, DETAIL)
    public ViewMarketingOrder getAllMoById(String moId) {
    	
        MarketingOrder marketingOrder = marketingOrderRepo.findByMoId(moId);
        
        if (marketingOrder == null) {
            throw new RuntimeException("Marketing Order not found for ID: " + moId);
        }

        List<HeaderMarketingOrder> headerMarketingOrders = headerMarketingOrderRepo.findByMoId(moId);
        List<DetailMarketingOrder> detailMarketingOrders = detailMarketingOrderRepo.findByMoId(moId);

        ViewMarketingOrder response = new ViewMarketingOrder();
        response.setMoId(marketingOrder.getMoId());
        response.setType(marketingOrder.getType());
        response.setDateValid(marketingOrder.getDateValid());
        
        System.out.println("halooo "+marketingOrder.getDateValid());
        response.setStatusFilled(marketingOrder.getStatusFilled());
        response.setStatus(marketingOrder.getStatus());
        response.setRevisionPpc(marketingOrder.getRevisionPpc());
        response.setRevisionMarketing(marketingOrder.getRevisionMarketing());

        List<ViewHeaderMarketingOrder> headerResponses = new ArrayList<>();
        for (HeaderMarketingOrder header : headerMarketingOrders) {
            ViewHeaderMarketingOrder headerResponse = new ViewHeaderMarketingOrder();
            
            headerResponse.setMoId(header.getMoId());
            headerResponse.setMonth(header.getMonth());
            headerResponse.setWdNormalTire(header.getWdNormalTire());
            headerResponse.setWdOtTl(header.getWdOtTl());
            headerResponse.setWdOtTt(header.getWdOtTt());
            headerResponse.setWdNormalTube(header.getWdNormalTube());
            headerResponse.setWdOtTube(header.getWdOtTube());
            headerResponse.setTotalWdTl(header.getTotalWdTl());
            headerResponse.setTotalWdTt(header.getTotalWdTt());
            headerResponse.setTotalWdTube(header.getTotalWdTube());
            headerResponse.setMaxCapTube(header.getMaxCapTube());
            headerResponse.setMaxCapTl(header.getMaxCapTl());
            headerResponse.setMaxCapTt(header.getMaxCapTt());
            headerResponse.setAirbagMachine(header.getAirbagMachine());
            headerResponse.setTl(header.getTl());
            headerResponse.setTt(header.getTt());
            headerResponse.setTotalMo(header.getTotalMo());
            headerResponse.setTlPercentage(header.getTlPercentage());
            headerResponse.setTtPercentage(header.getTtPercentage());
            headerResponse.setNoteOrderTl(header.getNoteOrderTl());

            
            headerResponses.add(headerResponse);
        }
        response.setDataHeaderMo(headerResponses);

        List<ViewDetailMarketingOrder> detailResponses = new ArrayList<>();
        for (DetailMarketingOrder detail : detailMarketingOrders) {
        	ViewDetailMarketingOrder detailResponse = new ViewDetailMarketingOrder();
        	
        	detailResponse.setDetailId(detail.getDetailId());
        	detailResponse.setMoId(detail.getMoId());
        	detailResponse.setCategory(detail.getCategory());
        	detailResponse.setPartNumber(detail.getPartNumber());
        	detailResponse.setDescription(detail.getDescription());
        	detailResponse.setMachineType(detail.getMachineType());
        	detailResponse.setCapacity(detail.getCapacity());
        	detailResponse.setQtyPerMould(detail.getQtyPerMould());
        	detailResponse.setQtyPerRak(detail.getQtyPerRak());
        	detailResponse.setMinOrder(detail.getMinOrder());
        	detailResponse.setMaxCapMonth0(detail.getMaxCapMonth0());
        	detailResponse.setMaxCapMonth1(detail.getMaxCapMonth1());
        	detailResponse.setMaxCapMonth2(detail.getMaxCapMonth2());
        	detailResponse.setInitialStock(detail.getInitialStock());
        	detailResponse.setSfMonth0(detail.getSfMonth0());
        	detailResponse.setSfMonth1(detail.getSfMonth1());
        	detailResponse.setSfMonth2(detail.getSfMonth2());
        	detailResponse.setMoMonth0(detail.getMoMonth0());
        	detailResponse.setMoMonth1(detail.getMoMonth1());
        	detailResponse.setMoMonth2(detail.getMoMonth2());
        	detailResponse.setPpd(detail.getPpd());
        	detailResponse.setCav(detail.getCav());
        	detailResponse.setLockStatusM0(detail.getLockStatusM0());
        	detailResponse.setLockStatusM1(detail.getLockStatusM1());
        	detailResponse.setLockStatusM2(detail.getLockStatusM2());
        	
        	List<Product> prodList = productRepo.findAll();

        	String itemCuring = null; 
        	for (Product product : prodList) {
        	    if (product.getPART_NUMBER().equals(detail.getPartNumber())) {
        	        itemCuring = product.getITEM_CURING(); 
        	        break;  
        	    }
        	}

        	detailResponse.setItemCuring(itemCuring);

            detailResponses.add(detailResponse);
        }

        response.setDataDetailMo(detailResponses);

        return response;
    }
    
    //UPDATE DETAIL MO (UPDATED BY ROLE MARKETING)
    public List<DetailMarketingOrder> updateDetailMOById(List<DetailMarketingOrder> detail) {
        List<DetailMarketingOrder> detailResponses = new ArrayList<>();
        List<ItemCuring> curingList = itemCuringRepo.findAll();
        
        String moId = detail.get(0).getMoId();
        MarketingOrder marketingOrder = marketingOrderRepo.findByMoId(moId);
        
        List<HeaderMarketingOrder> headerMOList = headerMarketingOrderRepo.findByMoId(moId);

        BigDecimal abM0 = BigDecimal.ZERO;
        BigDecimal abM1 = BigDecimal.ZERO;
        BigDecimal abM2 = BigDecimal.ZERO;
        BigDecimal tlM0 = BigDecimal.ZERO;
        BigDecimal tlM1 = BigDecimal.ZERO;
        BigDecimal tlM2 = BigDecimal.ZERO;
        BigDecimal ttM0 = BigDecimal.ZERO;
        BigDecimal ttM1 = BigDecimal.ZERO;
        BigDecimal ttM2 = BigDecimal.ZERO;
        String itemCuring = " ";
        
        System.out.println("header MO LIST EDIT " + headerMOList.size());
        System.out.println("haloooooooo " + detail.size());
        for (DetailMarketingOrder detaill : detail) {
        	
            DetailMarketingOrder detailMo = new DetailMarketingOrder();
            
            System.out.println("ini id detail " + detaill.getDetailId());
            detailMo.setDetailId(detaill.getDetailId());
            detailMo.setMoId(detaill.getMoId());
            detailMo.setDescription(detaill.getDescription());
            detailMo.setCategory(detaill.getCategory());
            detailMo.setMachineType(detaill.getMachineType());
            detailMo.setPartNumber(detaill.getPartNumber());
            detailMo.setCapacity(detaill.getCapacity());
            detailMo.setQtyPerMould(detaill.getQtyPerMould());
            detailMo.setQtyPerRak(detaill.getQtyPerRak());
            detailMo.setMinOrder(detaill.getMinOrder());
            detailMo.setMaxCapMonth0(detaill.getMaxCapMonth0());
            detailMo.setMaxCapMonth1(detaill.getMaxCapMonth1());
            detailMo.setMaxCapMonth2(detaill.getMaxCapMonth2());
            detailMo.setInitialStock(detaill.getInitialStock());

            detailMo.setSfMonth0(detaill.getSfMonth0());
            detailMo.setSfMonth1(detaill.getSfMonth1());
            detailMo.setSfMonth2(detaill.getSfMonth2());
            detailMo.setMoMonth0(detaill.getMoMonth0());
            detailMo.setMoMonth1(detaill.getMoMonth1());
            detailMo.setMoMonth2(detaill.getMoMonth2());
            detailMo.setLockStatusM0(detaill.getLockStatusM0());
            detailMo.setLockStatusM1(detaill.getLockStatusM1());
            detailMo.setLockStatusM2(detaill.getLockStatusM2());

            detailResponses.add(detailMo);
                        
            
            BigDecimal totalMO = detailMo.getMoMonth0(); // Total MO dari detail
            System.out.println("ini totalMO" + totalMO);

            BigDecimal prodType = BigDecimal.ZERO;
            BigDecimal hk = BigDecimal.ZERO;
            BigDecimal ppd = BigDecimal.ZERO;
            
            List<Product> prodList = productRepo.findAll();
            List<ProductType> prodTypeList = productTypeRepo.findAll();
            
	            // Menentukan PRODUCT_TYPE
	            for (Product pro : prodList) {
	                if (pro.getPART_NUMBER().equals(detailMo.getPartNumber())) {
	                    prodType = pro.getPRODUCT_TYPE_ID();
	                    itemCuring = pro.getITEM_CURING();
	                    break; 
	                }
	            }
	            
	        	BigDecimal HKTT = headerMOList.get(0).getTotalWdTt(); 
                BigDecimal HKTL = headerMOList.get(0).getTotalWdTl(); 
            	System.out.println("HK TT " + HKTT);
            	System.out.println("HK TL " + HKTL);


	            // Menentukan HK berdasarkan PRODUCT_TYPE
	            for (HeaderMarketingOrder headerMO : headerMOList) {
	  
	                headerMO.getTl();
	                
	
	                // Periksa PRODUCT_TYPE
	                for (ProductType prot : prodTypeList) {
	                    if (prot.getPRODUCT_TYPE_ID().equals(prodType)) {
	                        if (prot.getPRODUCT_TYPE().equals("TT")) {
	                            hk = HKTT; 
	                			ttM0 = ttM0.add(detailMo.getMoMonth0());
	                			ttM1 = ttM1.add(detailMo.getMoMonth1());
	                			ttM2 = ttM2.add(detailMo.getMoMonth2());
	                        } else if (prot.getPRODUCT_TYPE().equals("TL")) {
	                            hk = HKTL;
	                            tlM0 = tlM0.add(detailMo.getMoMonth0());
	                			tlM1 = tlM1.add(detailMo.getMoMonth1());
	                			tlM2 = tlM2.add(detailMo.getMoMonth2());
	                        }
	                        break;
	                    }
	                }
	
	                // Menghitung PPD berdasarkan HK
	                if (hk.compareTo(BigDecimal.ZERO) != 0) {
		                System.out.println("ini ppddd " + hk);
	                    ppd = totalMO.divide(hk, RoundingMode.HALF_UP); 
	                    
	                } else {
	                    System.err.println("Warning: HK is zero, setting ppd to zero.");
	                }
	                System.out.println("ini ppd " + ppd);
	
	                detailMo.setPpd(ppd);
	                BigDecimal cav = ppd.divide(detailMo.getCapacity(), RoundingMode.HALF_UP);
	                if (cav.compareTo(BigDecimal.ONE) < 0) {
		                detailMo.setCav(BigDecimal.ONE);
	                }else {
		                detailMo.setCav(cav);
	                }
	                
	                for(ItemCuring cur : curingList) {
	                	if(cur.getITEM_CURING().compareTo(itemCuring) == 0) {
	                		if(cur.getMACHINE_TYPE().compareTo("A/B") == 0) {
	                			abM0 = abM0.add(detailMo.getMoMonth0());
	                			abM1 = abM0.add(detailMo.getMoMonth1());
	                			abM2 = abM0.add(detailMo.getMoMonth2());
	                		}
	                	}
	                }
	                
		            headerMarketingOrderRepo.save(headerMO); 
	            }
	            detailResponses.add(detailMarketingOrderRepo.save(detailMo)); 
        	}
                	
        for (HeaderMarketingOrder hmo : headerMOList) {
        	
            if (hmo.getMonth().equals(marketingOrder.getMonth0())) {
                hmo.setAirbagMachine(abM0);
                hmo.setTl(tlM0);
                hmo.setTt(ttM0);
                hmo.setTotalMo(tlM0.add(ttM0));
                if (hmo.getTotalMo().compareTo(BigDecimal.ZERO) > 0) {
                    BigDecimal tlPercentage = hmo.getTl()
                        .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100))
                        .setScale(2, RoundingMode.HALF_UP);
                    hmo.setTlPercentage(tlPercentage);
                    BigDecimal ttPercentage = hmo.getTt()
                        .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100))
                        .setScale(2, RoundingMode.HALF_UP);
                    hmo.setTtPercentage(ttPercentage);
                } else {
                    hmo.setTlPercentage(BigDecimal.ZERO);
                    hmo.setTtPercentage(BigDecimal.ZERO);
                }
                
                headerMarketingOrderRepo.save(hmo);
                
            } else if (hmo.getMonth().equals(marketingOrder.getMonth1())) {
                hmo.setAirbagMachine(abM1);
                hmo.setTl(tlM1);
                hmo.setTt(ttM1);
                hmo.setTotalMo(tlM1.add(ttM1));
                if (hmo.getTotalMo().compareTo(BigDecimal.ZERO) > 0) {
                    BigDecimal tlPercentage = hmo.getTl()
                        .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100))
                        .setScale(2, RoundingMode.HALF_UP);
                    hmo.setTlPercentage(tlPercentage);
                    BigDecimal ttPercentage = hmo.getTt()
                        .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100))
                        .setScale(2, RoundingMode.HALF_UP);
                    hmo.setTtPercentage(ttPercentage);
                } else {
                    hmo.setTlPercentage(BigDecimal.ZERO);
                    hmo.setTtPercentage(BigDecimal.ZERO);
                }
                
                headerMarketingOrderRepo.save(hmo);
                
            } else if (hmo.getMonth().equals(marketingOrder.getMonth2())) {
                hmo.setAirbagMachine(abM2);
                hmo.setTl(tlM2);
                hmo.setTt(ttM2);
                hmo.setTotalMo(tlM2.add(ttM2));
                if (hmo.getTotalMo().compareTo(BigDecimal.ZERO) > 0) {
                    BigDecimal tlPercentage = hmo.getTl()
                        .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100))
                        .setScale(2, RoundingMode.HALF_UP);
                    hmo.setTlPercentage(tlPercentage);
                    
                    BigDecimal ttPercentage = hmo.getTt()
                        .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100))
                        .setScale(2, RoundingMode.HALF_UP);
                    hmo.setTtPercentage(ttPercentage);
                } else {
                    hmo.setTlPercentage(BigDecimal.ZERO);
                    hmo.setTtPercentage(BigDecimal.ZERO);
                }
            }
            headerMarketingOrderRepo.save(hmo);
            
            disableMarketingOrder(marketingOrder); 
           
            marketingOrder.setRevisionMarketing(BigDecimal.ZERO);
            
            marketingOrderRepo.save(marketingOrder);

        }
        	return detailResponses; 
    }
    
    //REVISION DETAIL MO (REVISION BY ROLE MARKETING)
    public int editMarketingOrderMarketing(EditMarketingOrderMarketing marketingOrderData) {
        
        List<DetailMarketingOrder> detail = marketingOrderData.getDetailMarketingOrder();
        List<HeaderMarketingOrder> headerMOList = marketingOrderData.getHeaderMarketingOrder();
        if (headerMOList.size() > 3) {
            headerMOList = headerMOList.subList(0, 3);
        }
        
        MarketingOrder marketingOrder = marketingOrderData.getMarketingOrder();
        marketingOrder.setStatus(BigDecimal.ONE);
        List<DetailMarketingOrder> detailResponses = new ArrayList<>();
        List<ItemCuring> curingList = itemCuringRepo.findAll();
        List<Product> prodList = productRepo.findAll();
        List<ProductType> prodTypeList = productTypeRepo.findAll();
        
        BigDecimal abM0 = BigDecimal.ZERO;
        BigDecimal abM1 = BigDecimal.ZERO;
        BigDecimal abM2 = BigDecimal.ZERO;
        BigDecimal tlM0 = BigDecimal.ZERO;
        BigDecimal tlM1 = BigDecimal.ZERO;
        BigDecimal tlM2 = BigDecimal.ZERO;
        BigDecimal ttM0 = BigDecimal.ZERO;
        BigDecimal ttM1 = BigDecimal.ZERO;
        BigDecimal ttM2 = BigDecimal.ZERO;
        String itemCuring = " ";
        
        for (DetailMarketingOrder detaill : detail) {
            DetailMarketingOrder detailMo = new DetailMarketingOrder();
            
            BigDecimal detailId = getNewDetailMarketingOrderId();
            System.out.println("ini id detailll " + detailId);

            String MOId = getLastIdMo();
            
            detailMo.setDetailId(detailId);
            detailMo.setMoId(detaill.getMoId());
            detailMo.setDescription(detaill.getDescription());
            detailMo.setCategory(detaill.getCategory());
            detailMo.setMachineType(detaill.getMachineType());
            detailMo.setPartNumber(detaill.getPartNumber());
            detailMo.setCapacity(detaill.getCapacity());
            detailMo.setQtyPerMould(detaill.getQtyPerMould());
            detailMo.setQtyPerRak(detaill.getQtyPerRak());
            detailMo.setMinOrder(detaill.getMinOrder());
            detailMo.setMaxCapMonth0(detaill.getMaxCapMonth0());
            detailMo.setMaxCapMonth1(detaill.getMaxCapMonth1());
            detailMo.setMaxCapMonth2(detaill.getMaxCapMonth2());
            detailMo.setInitialStock(detaill.getInitialStock());
            
            detailMo.setSfMonth0(detaill.getSfMonth0());
            detailMo.setSfMonth1(detaill.getSfMonth1());
            detailMo.setSfMonth2(detaill.getSfMonth2());
            detailMo.setMoMonth0(detaill.getMoMonth0());
            detailMo.setMoMonth1(detaill.getMoMonth1());
            detailMo.setMoMonth2(detaill.getMoMonth2());
            detailMo.setLockStatusM0(detaill.getLockStatusM0());
            detailMo.setLockStatusM1(detaill.getLockStatusM1());
            detailMo.setLockStatusM2(detaill.getLockStatusM2());

            detailResponses.add(detailMo);
            
            BigDecimal totalMO = detailMo.getMoMonth0();
            BigDecimal prodType = BigDecimal.ZERO;
            BigDecimal hk = BigDecimal.ZERO;
            BigDecimal ppd = BigDecimal.ZERO;
            
            for (Product pro : prodList) {
                System.out.println("ini " + 6);
                if (pro.getPART_NUMBER().equals(detailMo.getPartNumber())) {
                    prodType = pro.getPRODUCT_TYPE_ID();
                    itemCuring = pro.getITEM_CURING();
                    break; 
                }
            }
            
            BigDecimal HKTT = headerMOList.get(0).getTotalWdTt(); 
            BigDecimal HKTL = headerMOList.get(0).getTotalWdTl(); 
        	System.out.println("HK TT " + HKTT);
        	System.out.println("HK TL " + HKTL);
        	
        	for (ProductType prot : prodTypeList) {
                if (prot.getPRODUCT_TYPE_ID().equals(prodType)) {
                    if (prot.getPRODUCT_TYPE().equals("TT")) {
                        hk = HKTT;
                        ttM0 = ttM0.add(detaill.getMoMonth0());
                        ttM1 = ttM1.add(detaill.getMoMonth1());
                        ttM2 = ttM2.add(detaill.getMoMonth2());
                    } else if (prot.getPRODUCT_TYPE().equals("TL")) {
                        hk = HKTL;
                        tlM0 = tlM0.add(detaill.getMoMonth0());
                        tlM1 = tlM1.add(detaill.getMoMonth1());
                        tlM2 = tlM2.add(detaill.getMoMonth2());
                    }
                    break;
                }
            }
            
            // Hitung PPD
            if (hk.compareTo(BigDecimal.ZERO) != 0) {
                ppd = totalMO.divide(hk, RoundingMode.HALF_UP);
            } else {
                System.err.println("Warning: HK is zero, setting ppd to zero.");
                ppd = BigDecimal.ZERO;
            }
            
            detailMo.setPpd(ppd);
            
            // Hitung Cavity
            BigDecimal cav = ppd.divide(detaill.getCapacity(), RoundingMode.HALF_UP);
            detailMo.setCav(cav.compareTo(BigDecimal.ONE) < 0 ? BigDecimal.ONE : cav);
            
            // Hitung Airbag Machine
            for (ItemCuring cur : curingList) {
                if (cur.getITEM_CURING().compareTo(itemCuring) == 0) {
                    if (cur.getMACHINE_TYPE().compareTo("A/B") == 0) {
                        abM0 = abM0.add(detaill.getMoMonth0());
                        abM1 = abM1.add(detaill.getMoMonth1());
                        abM2 = abM2.add(detaill.getMoMonth2());
                    }
                }
            }
            
            detailResponses.add(detailMarketingOrderRepo.save(detailMo));
        }
        
        for (HeaderMarketingOrder hmo : headerMOList) {
        	BigDecimal newHeaderId = getNewHeaderMarketingOrderId();
        	hmo.setHeaderId(newHeaderId);
        	hmo.setStatus(BigDecimal.ONE);
        	
            System.out.println("ini " + 12);
            if (hmo.getMonth().equals(marketingOrder.getMonth0())) {
                hmo.setAirbagMachine(abM0);
                hmo.setTl(tlM0);
                hmo.setTt(ttM0);
                hmo.setTotalMo(tlM0.add(ttM0));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(marketingOrder.getMonth1())) {
                hmo.setAirbagMachine(abM1);
                hmo.setTl(tlM1);
                hmo.setTt(ttM1);
                hmo.setTotalMo(tlM1.add(ttM1));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(marketingOrder.getMonth2())) {
                hmo.setAirbagMachine(abM2);
                hmo.setTl(tlM2);
                hmo.setTt(ttM2);
                hmo.setTotalMo(tlM2.add(ttM2));
                updatePercentages(hmo);
            }
            headerMarketingOrderRepo.save(hmo);
        }
        System.out.println("ini MO" + marketingOrder.getMoId());

        if(marketingOrder != null) {
        	marketingOrder.setStatusFilled(BigDecimal.valueOf(3));
		}
		
        if (marketingOrder.getRevisionMarketing() == null) {
        	marketingOrder.setRevisionMarketing(BigDecimal.ONE); // Jika null atau 0, set ke 1
        } else {
        	marketingOrder.setRevisionMarketing(marketingOrder.getRevisionMarketing().add(BigDecimal.ONE)); // Tambah 1 pada revisi
        }
        	        
        marketingOrderRepo.save(marketingOrder);
        
        System.out.println("ini " + 13);
        return 1; 
    }
    
    //AR DEFECT REJECT (Role PPC MonthlyPlanning AR DEFECT REJECT)
    public int ArRejectDefectMO(EditMarketingOrderMarketing marketingOrderData) {
        
        List<DetailMarketingOrder> detail = marketingOrderData.getDetailMarketingOrder();
        List<HeaderMarketingOrder> headerMOList = marketingOrderData.getHeaderMarketingOrder();
        if (headerMOList.size() > 3) {
            headerMOList = headerMOList.subList(0, 3);
        }
        
        MarketingOrder marketingOrder = marketingOrderData.getMarketingOrder();
        marketingOrder.setStatus(BigDecimal.ONE);
        List<DetailMarketingOrder> detailResponses = new ArrayList<>();
        List<ItemCuring> curingList = itemCuringRepo.findAll();
        List<Product> prodList = productRepo.findAll();
        List<ProductType> prodTypeList = productTypeRepo.findAll();
        
        BigDecimal abM0 = BigDecimal.ZERO;
        BigDecimal abM1 = BigDecimal.ZERO;
        BigDecimal abM2 = BigDecimal.ZERO;
        BigDecimal tlM0 = BigDecimal.ZERO;
        BigDecimal tlM1 = BigDecimal.ZERO;
        BigDecimal tlM2 = BigDecimal.ZERO;
        BigDecimal ttM0 = BigDecimal.ZERO;
        BigDecimal ttM1 = BigDecimal.ZERO;
        BigDecimal ttM2 = BigDecimal.ZERO;
        String itemCuring = " ";
        
        for (DetailMarketingOrder detaill : detail) {
            DetailMarketingOrder detailMo = new DetailMarketingOrder();
            
            BigDecimal detailId = getNewDetailMarketingOrderId();
            System.out.println("ini id detailll " + detailId);

            String MOId = getLastIdMo();
            
            detailMo.setDetailId(detailId);
            detailMo.setMoId(detaill.getMoId());
            detailMo.setDescription(detaill.getDescription());
            detailMo.setCategory(detaill.getCategory());
            detailMo.setMachineType(detaill.getMachineType());
            detailMo.setPartNumber(detaill.getPartNumber());
            detailMo.setCapacity(detaill.getCapacity());
            detailMo.setQtyPerMould(detaill.getQtyPerMould());
            detailMo.setQtyPerRak(detaill.getQtyPerRak());
            detailMo.setMinOrder(detaill.getMinOrder());
            detailMo.setMaxCapMonth0(detaill.getMaxCapMonth0());
            detailMo.setMaxCapMonth1(detaill.getMaxCapMonth1());
            detailMo.setMaxCapMonth2(detaill.getMaxCapMonth2());
            detailMo.setInitialStock(detaill.getInitialStock());
            
            detailMo.setSfMonth0(detaill.getSfMonth0());
            detailMo.setSfMonth1(detaill.getSfMonth1());
            detailMo.setSfMonth2(detaill.getSfMonth2());
            detailMo.setMoMonth0(detaill.getMoMonth0());
            detailMo.setMoMonth1(detaill.getMoMonth1());
            detailMo.setMoMonth2(detaill.getMoMonth2());
            detailMo.setAr(detaill.getAr());
            detailMo.setDefect(detaill.getDefect());
            detailMo.setReject(detaill.getReject());
            detailMo.setLockStatusM0(detaill.getLockStatusM0());
            detailMo.setLockStatusM1(detaill.getLockStatusM1());
            detailMo.setLockStatusM2(detaill.getLockStatusM2());

            detailResponses.add(detailMo);
            
            BigDecimal totalMO = detailMo.getMoMonth0();
            BigDecimal prodType = BigDecimal.ZERO;
            BigDecimal hk = BigDecimal.ZERO;
            BigDecimal ppd = BigDecimal.ZERO;
            
            for (Product pro : prodList) {
                System.out.println("ini " + 6);
                if (pro.getPART_NUMBER().equals(detailMo.getPartNumber())) {
                    prodType = pro.getPRODUCT_TYPE_ID();
                    itemCuring = pro.getITEM_CURING();
                    break; 
                }
            }
            
            BigDecimal HKTT = headerMOList.get(0).getTotalWdTt(); 
            BigDecimal HKTL = headerMOList.get(0).getTotalWdTl(); 
        	System.out.println("HK TT " + HKTT);
        	System.out.println("HK TL " + HKTL);
        	
        	for (ProductType prot : prodTypeList) {
                if (prot.getPRODUCT_TYPE_ID().equals(prodType)) {
                    if (prot.getPRODUCT_TYPE().equals("TT")) {
                        hk = HKTT;
                        ttM0 = ttM0.add(detaill.getMoMonth0());
                        ttM1 = ttM1.add(detaill.getMoMonth1());
                        ttM2 = ttM2.add(detaill.getMoMonth2());
                    } else if (prot.getPRODUCT_TYPE().equals("TL")) {
                        hk = HKTL;
                        tlM0 = tlM0.add(detaill.getMoMonth0());
                        tlM1 = tlM1.add(detaill.getMoMonth1());
                        tlM2 = tlM2.add(detaill.getMoMonth2());
                    }
                    break;
                }
            }
            
            // Hitung PPD
            if (hk.compareTo(BigDecimal.ZERO) != 0) {
                ppd = totalMO.divide(hk, RoundingMode.HALF_UP);
            } else {
                System.err.println("Warning: HK is zero, setting ppd to zero.");
                ppd = BigDecimal.ZERO;
            }
            
            detailMo.setPpd(ppd);
            
            // Hitung Cavity
            BigDecimal cav = ppd.divide(detaill.getCapacity(), RoundingMode.HALF_UP);
            detailMo.setCav(cav.compareTo(BigDecimal.ONE) < 0 ? BigDecimal.ONE : cav);
            
            // Hitung Airbag Machine
            for (ItemCuring cur : curingList) {
                if (cur.getITEM_CURING().compareTo(itemCuring) == 0) {
                    if (cur.getMACHINE_TYPE().compareTo("A/B") == 0) {
                        abM0 = abM0.add(detaill.getMoMonth0());
                        abM1 = abM1.add(detaill.getMoMonth1());
                        abM2 = abM2.add(detaill.getMoMonth2());
                    }
                }
            }
            
            detailResponses.add(detailMarketingOrderRepo.save(detailMo));
        }
        
        for (HeaderMarketingOrder hmo : headerMOList) {
        	BigDecimal newHeaderId = getNewHeaderMarketingOrderId();
        	hmo.setHeaderId(newHeaderId);
        	hmo.setStatus(BigDecimal.ONE);
        	
            System.out.println("ini " + 12);
            if (hmo.getMonth().equals(marketingOrder.getMonth0())) {
                hmo.setAirbagMachine(abM0);
                hmo.setTl(tlM0);
                hmo.setTt(ttM0);
                hmo.setTotalMo(tlM0.add(ttM0));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(marketingOrder.getMonth1())) {
                hmo.setAirbagMachine(abM1);
                hmo.setTl(tlM1);
                hmo.setTt(ttM1);
                hmo.setTotalMo(tlM1.add(ttM1));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(marketingOrder.getMonth2())) {
                hmo.setAirbagMachine(abM2);
                hmo.setTl(tlM2);
                hmo.setTt(ttM2);
                hmo.setTotalMo(tlM2.add(ttM2));
                updatePercentages(hmo);
            }
            headerMarketingOrderRepo.save(hmo);
        }
        System.out.println("ini MO" + marketingOrder.getMoId());

        if(marketingOrder != null) {
        	marketingOrder.setStatusFilled(BigDecimal.valueOf(4));
		}
		
        if (marketingOrder.getRevisionMarketing() == null) {
        	marketingOrder.setRevisionMarketing(BigDecimal.ONE); // Jika null atau 0, set ke 1
        } else {
        	marketingOrder.setRevisionMarketing(marketingOrder.getRevisionMarketing().add(BigDecimal.ONE)); // Tambah 1 pada revisi
        }
        	        
        marketingOrderRepo.save(marketingOrder);
        
        System.out.println("ini " + 13);
        return 1; 
    }
	
	    private void updatePercentages(HeaderMarketingOrder hmo) {
	        if (hmo.getTotalMo().compareTo(BigDecimal.ZERO) > 0) {
	            BigDecimal tlPercentage = hmo.getTl()
	                .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
	                .multiply(BigDecimal.valueOf(100))
	                .setScale(2, RoundingMode.HALF_UP);
	            hmo.setTlPercentage(tlPercentage);
	            
	            BigDecimal ttPercentage = hmo.getTt()
	                .divide(hmo.getTotalMo(), 4, RoundingMode.HALF_UP)
	                .multiply(BigDecimal.valueOf(100))
	                .setScale(2, RoundingMode.HALF_UP);
	            hmo.setTtPercentage(ttPercentage);
	        } else {
	            hmo.setTlPercentage(BigDecimal.ZERO);
	            hmo.setTtPercentage(BigDecimal.ZERO);
	        }
	    }

    
    	//STATUS FILLED 1 (DI ISI PPC)
	    public void updateStatusFilledPpc(MarketingOrder marketingOrder) {
				MarketingOrder mo = marketingOrderRepo.findByMoId(marketingOrder.getMoId());
				if(mo != null) {
					mo.setStatusFilled(BigDecimal.valueOf(1));
				}
				marketingOrderRepo.save(mo);
		}
	    
	    //ENABLE 
	    public void enableMarketingOrder(MarketingOrder marketingOrder) {
			MarketingOrder mo = marketingOrderRepo.findByMoId(marketingOrder.getMoId());
			if(mo != null) {
				mo.setStatusFilled(BigDecimal.valueOf(2));
			}
			marketingOrderRepo.save(mo);
		}
	    
	    //DISABLE
	    public void disableMarketingOrder(MarketingOrder marketingOrder) {
	    		MarketingOrder mo = marketingOrderRepo.findByMoId(marketingOrder.getMoId());
	    		if(mo != null) {
	    			mo.setStatusFilled(BigDecimal.valueOf(3));
	    		}
	    		
	    		marketingOrderRepo.save(mo);
	    }
	    
	    public List<Map<String, Object>> getWorkDay(int month1, int year1, int month2, int year2, int month3, int year3) {
	        List<Map<String, Object>> headerList = new ArrayList<>();
	        
	        int[][] monthsAndYears = {
	            {month1, year1},
	            {month2, year2},
	            {month3, year3}
	        };

	        for (int[] monthYear : monthsAndYears) {
	            int month = monthYear[0];
	            int year = monthYear[1];
	            
	            Map<String, Object> result = headerMarketingOrderRepo.getMonthlyWorkData(month, year);
	            
	            Map<String, Object> headerMap = new HashMap<>();
                
	            headerMap.put("wdNormalTire", result.get("FINAL_WD") != null ? ((BigDecimal) result.get("FINAL_WD")).setScale(2, RoundingMode.HALF_UP) : BigDecimal.ZERO);
	            headerMap.put("wdOtTl", result.get("FINAL_OT_TL") != null ? ((BigDecimal) result.get("FINAL_OT_TL")).setScale(2, RoundingMode.HALF_UP) : BigDecimal.ZERO);
	            headerMap.put("wdOtTt", result.get("FINAL_OT_TT") != null ? ((BigDecimal) result.get("FINAL_OT_TT")).setScale(2, RoundingMode.HALF_UP) : BigDecimal.ZERO);
	            headerMap.put("wdNormalTube", result.get("FINAL_WD") != null ? ((BigDecimal) result.get("FINAL_WD")).setScale(2, RoundingMode.HALF_UP) : BigDecimal.ZERO); // Asumsi sama dengan FINAL_WD
	            headerMap.put("totalWdTl", result.get("TOTAL_OT_TL") != null ? ((BigDecimal) result.get("TOTAL_OT_TL")).setScale(2, RoundingMode.HALF_UP) : BigDecimal.ZERO);
	            headerMap.put("totalWdTt", result.get("TOTAL_OT_TT") != null ? ((BigDecimal) result.get("TOTAL_OT_TT")).setScale(2, RoundingMode.HALF_UP) : BigDecimal.ZERO);
                
                headerList.add(headerMap);
	        }
	        
	        return headerList;
	    }


	    
	    public ByteArrayInputStream exportMOExcel (String id) throws IOException {
	    	ByteArrayInputStream byteArrayInputStream = dataToExcel(id);
	    	return byteArrayInputStream;
	    }
	    
	    
	  //EXPORT MARKETING ORDER
	    public ByteArrayInputStream dataToExcel(String id) throws IOException {
	    	Optional<MarketingOrder> optionalMarketingOrder = marketingOrderRepo.findById(id);
	    	MarketingOrder marketingOrder = optionalMarketingOrder.get();
	    	List<HeaderMarketingOrder> headerMarketingOrder = headerMarketingOrderRepo.findByMoId(id);
	    	List<DetailMarketingOrder> detailMarketingOrder = detailMarketingOrderRepo.findByMoId(id);
	    	
	    	SimpleDateFormat monthFormat = new SimpleDateFormat("MMM");
	    		
	        String[] headerMO = {
	            "KETERANGAN", "HK Normal Tire", "HK Tube", "HK OT TL", "HK OT TT", "HK OT Tube",
	            "Total HK Tire TL", "Total HK Tire TT", "Total HK Tube" , "Kapasitas Maks Tube",
	            "Kapasitas Maks Tire TL", "Kapasitas Maks Tire TT",
	            "Kapasitas Mesin Airbag", "FED TL", "FED TT", "Total Marketing Order", 
	            "% FED TL", "% FED TT", "NOTE ORDER TL"
	        };

	        String[] detailMO = {
	            "Kategori", "Item", "Deskripsi", "Type Mesin", "Kapasitas 99,5%",
	            "Qty Mould", "Qty Per Rak", "Min Order", monthFormat.format(marketingOrder.getMonth0()),
	            monthFormat.format(marketingOrder.getMonth1()), monthFormat.format(marketingOrder.getMonth2()),
	            "Stok Awal", monthFormat.format(marketingOrder.getMonth0()), monthFormat.format(marketingOrder.getMonth1()),
	            monthFormat.format(marketingOrder.getMonth2()), monthFormat.format(marketingOrder.getMonth0()),
	            monthFormat.format(marketingOrder.getMonth1()), monthFormat.format(marketingOrder.getMonth2())
	        };

	        Workbook workbook = new XSSFWorkbook();
	        ByteArrayOutputStream out = new ByteArrayOutputStream();

	        try {
	            Sheet sheet = workbook.createSheet("Form MO");
	            
	            // Set column width
	            sheet.setColumnWidth(1, 2500);
	            sheet.setColumnWidth(2, 5000);
	            sheet.setColumnWidth(3, 9000);
	            sheet.setColumnWidth(4, 3000);
	            sheet.setColumnWidth(5, 4000);
	            sheet.setColumnWidth(6, 3000);
	            sheet.setColumnWidth(7, 3000);
	            sheet.setColumnWidth(8, 3000);
	            sheet.setColumnWidth(12, 3000);

	            // Font
	            Font candaraBold20 = workbook.createFont();
	            candaraBold20.setFontName("Candara");
	            candaraBold20.setFontHeightInPoints((short) 20);
	            candaraBold20.setBold(true);
	            
	            Font calibri11 = workbook.createFont();
	            calibri11.setFontName("Calibri");
	            calibri11.setFontHeightInPoints((short) 11);
	            
	            Font calibri12 = workbook.createFont();
	            calibri12.setFontName("Calibri");
	            calibri12.setFontHeightInPoints((short) 12);
	            
	            Font calibriBold11 = workbook.createFont();
	            calibriBold11.setFontName("Calibri");
	            calibriBold11.setFontHeightInPoints((short) 11);
	            calibriBold11.setBold(true);
	            
	            Font calibriBold12 = workbook.createFont();
	            calibriBold12.setFontName("Calibri");
	            calibriBold12.setFontHeightInPoints((short) 12);
	            calibriBold12.setBold(true);
	            // End Font
	            
	            // Background Color
	            XSSFColor lightBlueGray = new XSSFColor(new java.awt.Color(220, 230, 241), new DefaultIndexedColorMap());
	            XSSFColor lightGray = new XSSFColor(new java.awt.Color(217, 217, 217), new DefaultIndexedColorMap());
	            XSSFColor gold = new XSSFColor(new java.awt.Color(255, 242, 204), new DefaultIndexedColorMap());
	            XSSFColor lightOrange = new XSSFColor(new java.awt.Color(255, 192, 0), new DefaultIndexedColorMap());
	            
	            // Border cell style
	            CellStyle borderStyle = workbook.createCellStyle();
	            borderStyle.setBorderTop(BorderStyle.THIN);
	            borderStyle.setBorderBottom(BorderStyle.THIN);
	            borderStyle.setBorderLeft(BorderStyle.THIN);
	            borderStyle.setBorderRight(BorderStyle.THIN);
	            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
	            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
	            // End border cell style

	            // Style untuk cell
	            CellStyle candaraBold20NB = workbook.createCellStyle();
	            candaraBold20NB.setFont(candaraBold20);
	            candaraBold20NB.setAlignment(HorizontalAlignment.CENTER);
	            candaraBold20NB.setVerticalAlignment(VerticalAlignment.CENTER);
	            
	            CellStyle calibri11NB = workbook.createCellStyle();
	            calibri11NB.setFont(calibri11);
	            calibri11NB.setAlignment(HorizontalAlignment.LEFT);
	            calibri11NB.setVerticalAlignment(VerticalAlignment.CENTER);
	            
	            CellStyle calibriBold11B = workbook.createCellStyle();
	            calibriBold11B.cloneStyleFrom(borderStyle);
	            calibriBold11B.setFont(calibriBold11);
	            calibriBold11B.setAlignment(HorizontalAlignment.LEFT);
	            calibriBold11B.setVerticalAlignment(VerticalAlignment.CENTER);
	            
	            CellStyle calibriBold11BLightGrey = workbook.createCellStyle();
	            calibriBold11BLightGrey.cloneStyleFrom(borderStyle);
	            calibriBold11BLightGrey.setFont(calibriBold11);
	            calibriBold11BLightGrey.setAlignment(HorizontalAlignment.LEFT);
	            calibriBold11BLightGrey.setVerticalAlignment(VerticalAlignment.CENTER);
	            ((XSSFCellStyle) calibriBold11BLightGrey).setFillForegroundColor(lightGray);
	            calibriBold11BLightGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            
	            CellStyle calibriBold11BR = workbook.createCellStyle();
	            calibriBold11BR.cloneStyleFrom(borderStyle);
	            calibriBold11BR.setFont(calibriBold11);
	            calibriBold11BR.setAlignment(HorizontalAlignment.RIGHT);
	            calibriBold11BR.setVerticalAlignment(VerticalAlignment.CENTER);
	            
	            
	            
	            CellStyle calibri12B = workbook.createCellStyle();
	            calibri12B.cloneStyleFrom(borderStyle);
	            calibri12B.setFont(calibri12);
	            calibri12B.setAlignment(HorizontalAlignment.LEFT);
	            calibri12B.setVerticalAlignment(VerticalAlignment.CENTER);
	            
	            CellStyle calibriBold12BL = workbook.createCellStyle();
	            calibriBold12BL.cloneStyleFrom(borderStyle);
	            calibriBold12BL.setFont(calibriBold12);
	            calibriBold12BL.setAlignment(HorizontalAlignment.LEFT);
	            calibriBold12BL.setVerticalAlignment(VerticalAlignment.CENTER);
	            ((XSSFCellStyle) calibriBold12BL).setFillForegroundColor(lightBlueGray);
	            calibriBold12BL.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            
	            CellStyle calibriBold12BC = workbook.createCellStyle();
	            calibriBold12BC.cloneStyleFrom(borderStyle);
	            calibriBold12BC.setFont(calibriBold12);
	            calibriBold12BC.setAlignment(HorizontalAlignment.CENTER);
	            calibriBold12BC.setVerticalAlignment(VerticalAlignment.CENTER);
	            ((XSSFCellStyle) calibriBold12BC).setFillForegroundColor(lightBlueGray);
	            calibriBold12BC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            
	            DataFormat dataFormat = workbook.createDataFormat();
	            CellStyle calibri12Num = workbook.createCellStyle();
	            calibri12Num.cloneStyleFrom(borderStyle);
	            calibri12Num.setFont(calibri12);
	            calibri12Num.setAlignment(HorizontalAlignment.RIGHT);
	            calibri12Num.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri12Num.setDataFormat(dataFormat.getFormat("#,##"));
	            
	            CellStyle calibri12NumLightGrey = workbook.createCellStyle();
	            calibri12NumLightGrey.cloneStyleFrom(borderStyle);
	            calibri12NumLightGrey.setFont(calibri12);
	            calibri12NumLightGrey.setAlignment(HorizontalAlignment.RIGHT);
	            calibri12NumLightGrey.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri12NumLightGrey.setDataFormat(dataFormat.getFormat("#,##"));
	            ((XSSFCellStyle) calibri12NumLightGrey).setFillForegroundColor(lightGray);
	            calibri12NumLightGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            
	            CellStyle calibri12NumGold = workbook.createCellStyle();
	            calibri12NumGold.cloneStyleFrom(borderStyle);
	            calibri12NumGold.setFont(calibri12);
	            calibri12NumGold.setAlignment(HorizontalAlignment.RIGHT);
	            calibri12NumGold.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri12NumGold.setDataFormat(dataFormat.getFormat("#,##"));
	            ((XSSFCellStyle) calibri12NumGold).setFillForegroundColor(gold);
	            calibri12NumGold.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            
	            CellStyle calibri12NumLightOrange = workbook.createCellStyle();
	            calibri12NumLightOrange.cloneStyleFrom(borderStyle);
	            calibri12NumLightOrange.setFont(calibri12);
	            calibri12NumLightOrange.setAlignment(HorizontalAlignment.RIGHT);
	            calibri12NumLightOrange.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri12NumLightOrange.setDataFormat(dataFormat.getFormat("#,##"));
	            ((XSSFCellStyle) calibri12NumLightOrange).setFillForegroundColor(lightOrange);
	            calibri12NumLightOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            
	            CellStyle calibri11NumBold = workbook.createCellStyle();
	            calibri11NumBold.cloneStyleFrom(borderStyle);
	            calibri11NumBold.setFont(calibriBold11);
	            calibri11NumBold.setAlignment(HorizontalAlignment.RIGHT);
	            calibri11NumBold.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri11NumBold.setDataFormat(dataFormat.getFormat("#,##"));
	            
	            CellStyle calibri11DecBold = workbook.createCellStyle();
	            calibri11DecBold.cloneStyleFrom(borderStyle);
	            calibri11DecBold.setFont(calibriBold11);
	            calibri11DecBold.setAlignment(HorizontalAlignment.RIGHT);
	            calibri11DecBold.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri11DecBold.setDataFormat(dataFormat.getFormat("#,##0.00"));
	            
	            CellStyle calibri11PerBold = workbook.createCellStyle();
	            calibri11PerBold.cloneStyleFrom(borderStyle);
	            calibri11PerBold.setFont(calibriBold11);
	            calibri11PerBold.setAlignment(HorizontalAlignment.RIGHT);
	            calibri11PerBold.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri11PerBold.setDataFormat(dataFormat.getFormat("#,##0.00\"%\""));
	            
	            CellStyle calibri11NumBoldLightGrey = workbook.createCellStyle();
	            calibri11NumBoldLightGrey.cloneStyleFrom(borderStyle);
	            calibri11NumBoldLightGrey.setFont(calibriBold11);
	            calibri11NumBoldLightGrey.setAlignment(HorizontalAlignment.RIGHT);
	            calibri11NumBoldLightGrey.setVerticalAlignment(VerticalAlignment.CENTER);
	            calibri11NumBoldLightGrey.setDataFormat(dataFormat.getFormat("#,##"));
	            ((XSSFCellStyle) calibri11NumBoldLightGrey).setFillForegroundColor(lightGray);
	            calibri11NumBoldLightGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            //end style

	            // Header
	            // Title
	            sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 11));
	            Row titleRow = sheet.createRow(0);
	            Cell titleCell = titleRow.createCell(1);
	            titleCell.setCellValue("Form Input Marketing Order");
	            titleCell.setCellStyle(candaraBold20NB);
	            
	            // Month
	            sheet.addMergedRegion(new CellRangeAddress(2, 4, 1, 11));
	            Row monthRow = sheet.createRow(2);
	            Cell monthCell = monthRow.createCell(1);
	            monthCell.setCellStyle(candaraBold20NB);
	            monthCell.setCellValue(monthFormat.format(marketingOrder.getMonth0()) + " - " + monthFormat.format(marketingOrder.getMonth1()) + " - " + monthFormat.format(marketingOrder.getMonth2()));

	            // Revision
	            Row revisionRow = sheet.createRow(13);
	            Cell revisionCell = revisionRow.createCell(2);
	            revisionCell.setCellStyle(calibri11NB);
	            revisionCell.setCellValue("Revisi :");
	            
	            
	            revisionCell = revisionRow.createCell(3);
	            revisionCell.setCellStyle(calibri11NB);
	            revisionCell.setCellValue(marketingOrder.getRevisionPpc().toString());

	            // Date
	            Row dateRow = sheet.createRow(14);
	            Cell dateCell = dateRow.createCell(2);
	            dateCell.setCellStyle(calibri11NB);
	            dateCell.setCellValue("Tanggal :");
	            
	            dateCell = dateRow.createCell(3);
	            dateCell.setCellStyle(calibri11NB);
	            Date dateValue = marketingOrder.getDateValid();
		        SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMM yyyy");
		        String formattedDate = dateFormat.format(dateValue);
		        dateCell.setCellValue(formattedDate);
	            
	            //MO Header
	            Row headerMORow = sheet.createRow(1);
	            Cell headerMOCell = headerMORow.createCell(13);
	            for (int i=0;i<headerMO.length;i++) {
	            	if (i == 0) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	                    titleCell = titleRow.createCell(13);
	                    titleCell.setCellValue(headerMO[i]);
	                    titleCell.setCellStyle(calibriBold12BL);
	                    titleCell = titleRow.createCell(14);
	                    titleCell.setCellStyle(calibriBold12BL);
	                    titleCell = titleRow.createCell(15);
	                    titleCell.setCellStyle(calibriBold12BL);
	                    
	                    titleCell = titleRow.createCell(16);
	                    titleCell.setCellValue(monthFormat.format(marketingOrder.getMonth0()));
	                    titleCell.setCellStyle(calibriBold12BL);
	                    titleCell = titleRow.createCell(17);
	                    titleCell.setCellValue(monthFormat.format(marketingOrder.getMonth1()));
	                    titleCell.setCellStyle(calibriBold12BL);
	                    titleCell = titleRow.createCell(18);
	                    titleCell.setCellValue(monthFormat.format(marketingOrder.getMonth2()));
	                    titleCell.setCellStyle(calibriBold12BL);
	            	} else if (i == 1) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getWdNormalTire() != null ? headerMarketingOrder.get(0).getWdNormalTire().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getWdNormalTire() != null ? headerMarketingOrder.get(1).getWdNormalTire().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getWdNormalTire() != null ? headerMarketingOrder.get(2).getWdNormalTire().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 2) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		monthCell = monthRow.createCell(13);
	            		monthCell.setCellValue(headerMO[i]);
	            		monthCell.setCellStyle(calibriBold11B);
	            		monthCell = monthRow.createCell(14);
	            		monthCell.setCellStyle(calibriBold11B);
	            		monthCell = monthRow.createCell(15);
	            		monthCell.setCellStyle(calibriBold11B);
	                    
	            		monthCell = monthRow.createCell(16);
	            		monthCell.setCellValue(headerMarketingOrder.get(0).getWdNormalTube() != null ? headerMarketingOrder.get(0).getWdNormalTube().doubleValue() : 0);
	            		monthCell.setCellStyle(calibri11DecBold);
	            		monthCell = monthRow.createCell(17);
	            		monthCell.setCellValue(headerMarketingOrder.get(1).getWdNormalTube() != null ? headerMarketingOrder.get(1).getWdNormalTube().doubleValue() : 0);
	            		monthCell.setCellStyle(calibri11DecBold);
	            		monthCell = monthRow.createCell(18);
	            		monthCell.setCellValue(headerMarketingOrder.get(2).getWdNormalTube() != null ? headerMarketingOrder.get(2).getWdNormalTube().doubleValue() : 0);
	            		monthCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 3) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	            		headerMOCell = headerMORow.createCell(13);
	            		headerMOCell.setCellValue(headerMO[i]);
	            		headerMOCell.setCellStyle(calibriBold11B);
	            		headerMOCell = headerMORow.createCell(14);
	            		headerMOCell.setCellStyle(calibriBold11B);
	            		headerMOCell = headerMORow.createCell(15);
	            		headerMOCell.setCellStyle(calibriBold11B);
	                    
	            		headerMOCell = headerMORow.createCell(16);
	            		headerMOCell.setCellValue(headerMarketingOrder.get(0).getWdOtTl() != null ? headerMarketingOrder.get(0).getWdOtTl().doubleValue() : 0);
	            		headerMOCell.setCellStyle(calibri11DecBold);
	            		headerMOCell = headerMORow.createCell(17);
	            		headerMOCell.setCellValue(headerMarketingOrder.get(1).getWdOtTl() != null ? headerMarketingOrder.get(1).getWdOtTl().doubleValue() : 0);
	            		headerMOCell.setCellStyle(calibri11DecBold);
	            		headerMOCell = headerMORow.createCell(18);
	            		headerMOCell.setCellValue(headerMarketingOrder.get(2).getWdOtTl() != null ? headerMarketingOrder.get(2).getWdOtTl().doubleValue() : 0);
	            		headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 4) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getWdOtTt() != null ? headerMarketingOrder.get(0).getWdOtTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getWdOtTt() != null ? headerMarketingOrder.get(1).getWdOtTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getWdOtTt() != null ? headerMarketingOrder.get(2).getWdOtTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 5) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getWdOtTube() != null ? headerMarketingOrder.get(0).getWdOtTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getWdOtTube() != null ? headerMarketingOrder.get(1).getWdOtTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getWdOtTube() != null ? headerMarketingOrder.get(2).getWdOtTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 6) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getTotalWdTl() != null ? headerMarketingOrder.get(0).getTotalWdTl().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getTotalWdTl() != null ? headerMarketingOrder.get(1).getTotalWdTl().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getTotalWdTl() != null ? headerMarketingOrder.get(2).getTotalWdTl().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 7) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getTotalWdTt() != null ? headerMarketingOrder.get(0).getTotalWdTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getTotalWdTt() != null ? headerMarketingOrder.get(1).getTotalWdTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getTotalWdTt() != null ? headerMarketingOrder.get(2).getTotalWdTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 8) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getTotalWdTube() != null ? headerMarketingOrder.get(0).getTotalWdTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getTotalWdTube() != null ? headerMarketingOrder.get(1).getTotalWdTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getTotalWdTube() != null ? headerMarketingOrder.get(2).getTotalWdTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11DecBold);
	            	} else if (i == 9) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getMaxCapTube() != null ? headerMarketingOrder.get(0).getMaxCapTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getMaxCapTube() != null ? headerMarketingOrder.get(1).getMaxCapTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getMaxCapTube() != null ? headerMarketingOrder.get(2).getMaxCapTube().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	            	} else if (i == 10) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getMaxCapTl() != null ? headerMarketingOrder.get(0).getMaxCapTl().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getMaxCapTl() != null ? headerMarketingOrder.get(1).getMaxCapTl().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getMaxCapTl() != null ? headerMarketingOrder.get(2).getMaxCapTl().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	            	} else if (i == 11) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getMaxCapTt() != null ? headerMarketingOrder.get(0).getMaxCapTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getMaxCapTt() != null ? headerMarketingOrder.get(1).getMaxCapTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getMaxCapTt() != null ? headerMarketingOrder.get(2).getMaxCapTt().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	            	} else if (i == 12) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getAirbagMachine() != null ? headerMarketingOrder.get(0).getAirbagMachine().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getAirbagMachine() != null ? headerMarketingOrder.get(1).getAirbagMachine().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getAirbagMachine() != null ? headerMarketingOrder.get(2).getAirbagMachine().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11NumBold);
	            	} else if (i == 13) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		revisionCell = revisionRow.createCell(13);
	            		revisionCell.setCellValue(headerMO[i]);
	            		revisionCell.setCellStyle(calibriBold11B);
	            		revisionCell = revisionRow.createCell(14);
	            		revisionCell.setCellStyle(calibriBold11B);
	            		revisionCell = revisionRow.createCell(15);
	            		revisionCell.setCellStyle(calibriBold11B);
	                    
	            		revisionCell = revisionRow.createCell(16);
	            		revisionCell.setCellValue(headerMarketingOrder.get(0).getTl() != null ? headerMarketingOrder.get(0).getTl().doubleValue() : 0);
	            		revisionCell.setCellStyle(calibri11NumBold);
	            		revisionCell = revisionRow.createCell(17);
	            		revisionCell.setCellValue(headerMarketingOrder.get(1).getTl() != null ? headerMarketingOrder.get(1).getTl().doubleValue() : 0);
	            		revisionCell.setCellStyle(calibri11NumBold);
	            		revisionCell = revisionRow.createCell(18);
	            		revisionCell.setCellValue(headerMarketingOrder.get(2).getTl() != null ? headerMarketingOrder.get(2).getTl().doubleValue() : 0);
	            		revisionCell.setCellStyle(calibri11NumBold);
	            	} else if (i == 14) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		dateCell = dateRow.createCell(13);
	            		dateCell = dateRow.createCell(13);
	            		dateCell.setCellValue(headerMO[i]);
	            		dateCell.setCellStyle(calibriBold11B);
	            		dateCell = dateRow.createCell(14);
	            		dateCell.setCellStyle(calibriBold11B);
	            		dateCell = dateRow.createCell(15);
	            		dateCell.setCellStyle(calibriBold11B);
	                    
	            		dateCell = dateRow.createCell(16);
	            		dateCell.setCellValue(headerMarketingOrder.get(0).getTt() != null ? headerMarketingOrder.get(0).getTt().doubleValue() : 0);
	            		dateCell.setCellStyle(calibri11NumBold);
	            		dateCell = dateRow.createCell(17);
	            		dateCell.setCellValue(headerMarketingOrder.get(1).getTt() != null ? headerMarketingOrder.get(1).getTt().doubleValue() : 0);
	            		dateCell.setCellStyle(calibri11NumBold);
	            		dateCell = dateRow.createCell(18);
	            		dateCell.setCellValue(headerMarketingOrder.get(2).getTt() != null ? headerMarketingOrder.get(2).getTt().doubleValue() : 0);
	            		dateCell.setCellStyle(calibri11NumBold);
	            	} else if (i == 15) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	            		headerMOCell = headerMORow.createCell(13);
	            		headerMOCell.setCellValue(headerMO[i]);
	            		headerMOCell.setCellStyle(calibriBold11BLightGrey);
	            		headerMOCell = headerMORow.createCell(14);
	            		headerMOCell.setCellStyle(calibriBold11BLightGrey);
	            		headerMOCell = headerMORow.createCell(15);
	            		headerMOCell.setCellStyle(calibriBold11BLightGrey);
	                    
	            		headerMOCell = headerMORow.createCell(16);
	            		headerMOCell.setCellValue(headerMarketingOrder.get(0).getTotalMo() != null ? headerMarketingOrder.get(0).getTotalMo().doubleValue() : 0);
	            		headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	            		headerMOCell = headerMORow.createCell(17);
	            		headerMOCell.setCellValue(headerMarketingOrder.get(1).getTotalMo() != null ? headerMarketingOrder.get(1).getTotalMo().doubleValue() : 0);
	            		headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	            		headerMOCell = headerMORow.createCell(18);
	            		headerMOCell.setCellValue(headerMarketingOrder.get(2).getTotalMo() != null ? headerMarketingOrder.get(2).getTotalMo().doubleValue() : 0);
	            		headerMOCell.setCellStyle(calibri11NumBoldLightGrey);
	            	} else if (i == 16) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getTlPercentage() != null ? headerMarketingOrder.get(0).getTlPercentage().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11PerBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getTlPercentage() != null ? headerMarketingOrder.get(1).getTlPercentage().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11PerBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getTlPercentage() != null ? headerMarketingOrder.get(2).getTlPercentage().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11PerBold);
	            	} else if (i == 17) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getTtPercentage() != null ? headerMarketingOrder.get(0).getTtPercentage().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11PerBold);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getTtPercentage() != null ? headerMarketingOrder.get(1).getTtPercentage().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11PerBold);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getTtPercentage() != null ? headerMarketingOrder.get(2).getTtPercentage().doubleValue() : 0);
	                    headerMOCell.setCellStyle(calibri11PerBold);
	            	} else if (i == 18) {
	            		sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 15));
	            		headerMORow = sheet.createRow(i);
	                    headerMOCell = headerMORow.createCell(13);
	                    headerMOCell.setCellValue(headerMO[i]);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(14);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(15);
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    
	                    headerMOCell = headerMORow.createCell(16);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(0).getNoteOrderTl() != null ? headerMarketingOrder.get(0).getNoteOrderTl() : "");
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(17);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(1).getNoteOrderTl() != null ? headerMarketingOrder.get(1).getNoteOrderTl() : "");
	                    headerMOCell.setCellStyle(calibriBold11B);
	                    headerMOCell = headerMORow.createCell(18);
	                    headerMOCell.setCellValue(headerMarketingOrder.get(2).getNoteOrderTl() != null ? headerMarketingOrder.get(2).getNoteOrderTl() : "");
	                    headerMOCell.setCellStyle(calibriBold11B);
	            	}
	            }
	            
	            Row rowHeaderDetail1 = sheet.createRow(19);
	            Row rowHeaderDetail2 = sheet.createRow(20);
	            for (int i=0;i<detailMO.length;i++) {
	            	if (i == 8 || i == 9 || i == 10 || i == 12 || i == 13 || i == 14 || i == 15 || i == 16 || i == 17) {
	            		if (i == 8 || i == 12 || i == 15) {
	            			sheet.addMergedRegion(new CellRangeAddress(19, 19, i+1, i+3));
	            			Cell cellHeaderDetail = rowHeaderDetail1.createCell(i+1);
	            			if (i == 8) {
	            				cellHeaderDetail.setCellValue("Kapasitas Maksimum");
	            			} else if (i == 12) {
	            				cellHeaderDetail.setCellValue("Sales Forecast");
	            			} else if (i == 15) {
	            				cellHeaderDetail.setCellValue("Marketing Order");
	            			}
	                        cellHeaderDetail.setCellStyle(calibriBold12BC);
	                        cellHeaderDetail = rowHeaderDetail1.createCell(i+2);
	                        cellHeaderDetail.setCellStyle(calibriBold12BC);
	                        cellHeaderDetail = rowHeaderDetail1.createCell(i+3);
	                        cellHeaderDetail.setCellStyle(calibriBold12BC);
	            		}
	            		Cell cellHeaderDetail = rowHeaderDetail2.createCell(i+1);
	                    cellHeaderDetail.setCellValue(detailMO[i]);
	                    cellHeaderDetail.setCellStyle(calibriBold12BC);
	            	} else {
	            		sheet.addMergedRegion(new CellRangeAddress(19, 20, i+1, i+1));
	                    
	                    Cell cellHeaderDetail = rowHeaderDetail1.createCell(i+1);
	                    cellHeaderDetail.setCellValue(detailMO[i]);
	                    cellHeaderDetail.setCellStyle(calibriBold12BC);
	                    
	                    cellHeaderDetail = rowHeaderDetail2.createCell(i+1);
	                    cellHeaderDetail.setCellStyle(calibriBold12BC);
	            	}
	            }
	            
	            // Mengisi data
	            int rowIndex = 21;
	            for (DetailMarketingOrder dMo : detailMarketingOrder) {
	                Row dataRow = sheet.createRow(rowIndex++);

	                Cell categoryCell = dataRow.createCell(1);
	                categoryCell.setCellValue(dMo.getCategory() != null ?  String.valueOf(dMo.getCategory()) : "");
	                categoryCell.setCellStyle(calibri12B);
	                
	                Cell itemCell = dataRow.createCell(2);
	                itemCell.setCellValue(dMo.getPartNumber() != null ? String.valueOf(dMo.getPartNumber()) : "");
	                itemCell.setCellStyle(calibri12B);

	                Cell descriptionCell = dataRow.createCell(3);
	                descriptionCell.setCellValue(dMo.getDescription() != null ? String.valueOf(dMo.getDescription()) : "");
	                descriptionCell.setCellStyle(calibri12B);

	                Cell machineCell = dataRow.createCell(4);
	                machineCell.setCellValue(dMo.getMachineType() != null ? String.valueOf(dMo.getMachineType()) : "");
	                machineCell.setCellStyle(calibri12B);

	                Cell capCell = dataRow.createCell(5);
	                capCell.setCellValue(dMo.getCapacity() != null ? ((Number) dMo.getCapacity()).doubleValue() : 0);
	                capCell.setCellStyle(calibri12Num);

	                Cell qtyMouldCell = dataRow.createCell(6);
	                qtyMouldCell.setCellValue(dMo.getQtyPerMould() != null ? ((Number) dMo.getQtyPerMould()).doubleValue() : 0);
	                qtyMouldCell.setCellStyle(calibri12Num);

	                Cell qtyPerRakCell = dataRow.createCell(7);
	                qtyPerRakCell.setCellValue(dMo.getQtyPerRak() != null ? ((Number) dMo.getQtyPerRak()).doubleValue() : 0);
	                qtyPerRakCell.setCellStyle(calibri12Num);

	                Cell minOrderCell = dataRow.createCell(8);
	                minOrderCell.setCellValue(dMo.getMinOrder() != null ? ((Number) dMo.getMinOrder()).doubleValue() : 0);
	                minOrderCell.setCellStyle(calibri12Num);

	                Cell maxCap0Cell = dataRow.createCell(9);
	                maxCap0Cell.setCellValue(dMo.getMaxCapMonth0() != null ? ((Number) dMo.getMaxCapMonth0()).doubleValue() : 0);
	                maxCap0Cell.setCellStyle(calibri12Num);

	                Cell maxCap1Cell = dataRow.createCell(10);
	                maxCap1Cell.setCellValue(dMo.getMaxCapMonth1() != null ? ((Number) dMo.getMaxCapMonth1()).doubleValue() : 0);
	                maxCap1Cell.setCellStyle(calibri12Num);

	                Cell maxCap2Cell = dataRow.createCell(11);
	                maxCap2Cell.setCellValue(dMo.getMaxCapMonth2() != null ? ((Number) dMo.getMaxCapMonth2()).doubleValue() : 0);
	                maxCap2Cell.setCellStyle(calibri12Num);

	                Cell initStockCell = dataRow.createCell(12);
	                initStockCell.setCellValue(dMo.getInitialStock() != null ? ((Number) dMo.getInitialStock()).doubleValue() : 0);
	                initStockCell.setCellStyle(calibri12NumLightGrey);

	                Cell salesForecast0Cell = dataRow.createCell(13);
	                salesForecast0Cell.setCellValue(dMo.getSfMonth0() != null ? ((Number) dMo.getSfMonth0()).doubleValue() : 0);
	                salesForecast0Cell.setCellStyle(calibri12NumGold);

	                Cell salesForecast1Cell = dataRow.createCell(14);
	                salesForecast1Cell.setCellValue(dMo.getSfMonth1() != null ? ((Number) dMo.getSfMonth1()).doubleValue() : 0);
	                salesForecast1Cell.setCellStyle(calibri12NumGold);

	                Cell salesForecast2Cell = dataRow.createCell(15);
	                salesForecast2Cell.setCellValue(dMo.getSfMonth2() != null ? ((Number) dMo.getSfMonth2()).doubleValue() : 0);
	                salesForecast2Cell.setCellStyle(calibri12NumGold);

	                Cell marketingOrder0Cell = dataRow.createCell(16);
	                marketingOrder0Cell.setCellValue(dMo.getMoMonth0() != null ? ((Number) dMo.getMoMonth0()).doubleValue() : 0);
	                marketingOrder0Cell.setCellStyle(calibri12NumLightOrange);

	                Cell marketingOrder1Cell = dataRow.createCell(17);
	                marketingOrder1Cell.setCellValue(dMo.getMoMonth1() != null ? ((Number) dMo.getMoMonth1()).doubleValue() : 0);
	                marketingOrder1Cell.setCellStyle(calibri12NumLightOrange);

	                Cell marketingOrder2Cell = dataRow.createCell(18);
	                marketingOrder2Cell.setCellValue(dMo.getMoMonth2() != null ? ((Number) dMo.getMoMonth2()).doubleValue() : 0);
	                marketingOrder2Cell.setCellStyle(calibri12NumLightOrange);
	            }
	            
	            workbook.write(out); // Menulis data ke output stream
	            return new ByteArrayInputStream(out.toByteArray());
	        } catch (IOException e) {
	            e.printStackTrace();
	            System.out.println("Fail to export data");
	            return null;
	        } finally {
	            out.close(); // Tutup output stream setelah selesai
	        }
	    }

	}

	
