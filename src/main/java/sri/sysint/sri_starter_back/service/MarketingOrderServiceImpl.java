package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
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
import sri.sysint.sri_starter_back.repository.MonthlyPlanRepo;
import sri.sysint.sri_starter_back.repository.HeaderMarketingOrderRepo;
import sri.sysint.sri_starter_back.repository.ItemCuringRepo;
import sri.sysint.sri_starter_back.repository.ProductRepo;
import sri.sysint.sri_starter_back.repository.ProductTypeRepo;
import sri.sysint.sri_starter_back.repository.SettingRepo;
import sri.sysint.sri_starter_back.model.HeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.ItemCuring;
import sri.sysint.sri_starter_back.model.MarketingOrder;
import sri.sysint.sri_starter_back.model.MonthlyPlan;
import sri.sysint.sri_starter_back.model.MonthlyPlanningCuring;
import sri.sysint.sri_starter_back.model.DetailDailyMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.DetailMarketingOrder;
import sri.sysint.sri_starter_back.model.DetailMonthlyPlanCuring;
import sri.sysint.sri_starter_back.model.Product;
import sri.sysint.sri_starter_back.model.ProductType;
import sri.sysint.sri_starter_back.model.WorkDay;
import sri.sysint.sri_starter_back.model.transaksi.EditMarketingOrderMarketing;
import sri.sysint.sri_starter_back.model.transaksi.GetAllTypeMarketingOrder;
import sri.sysint.sri_starter_back.model.transaksi.SaveFinalMarketingOrder;
import sri.sysint.sri_starter_back.model.transaksi.SaveMarketingOrderPPC;
import sri.sysint.sri_starter_back.model.transaksi.ViewMonthlyPlanning;
import sri.sysint.sri_starter_back.model.view.ViewDetailMarketingOrder;
import sri.sysint.sri_starter_back.model.view.ViewDetailShiftMonthlyPlan;
import sri.sysint.sri_starter_back.model.view.ViewDistinctMarketingOrder;
import sri.sysint.sri_starter_back.model.view.ViewHeaderMarketingOrder;
import sri.sysint.sri_starter_back.model.view.ViewMachineCuring;
import sri.sysint.sri_starter_back.model.view.ViewMarketingOrder;
import sri.sysint.sri_starter_back.repository.CTCuringRepo;
import sri.sysint.sri_starter_back.repository.DetailMarketingOrderRepo;


@Service
@Transactional
public class MarketingOrderServiceImpl {
	
	@Autowired
    private MarketingOrderRepo marketingOrderRepo;
	
	@Autowired
    private MonthlyPlanRepo monthlyPlanningRepo;
	
	@Autowired
    private HeaderMarketingOrderRepo headerMarketingOrderRepo;
	
	@Autowired
    private DetailMarketingOrderRepo detailMarketingOrderRepo;
	
	@Autowired
    private ItemCuringRepo itemCuringRepo;
	
	@Autowired
    private SettingRepo settingRepo;
	
	@Autowired
    private ProductRepo productRepo;
	
	@Autowired
    private ProductTypeRepo productTypeRepo;
	
	@Autowired
    private CTCuringRepo ctCuringRepo;
	
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
    
    //Find Distinct Month
    public List<ViewDistinctMarketingOrder> findDistinctMonths(){
        List<Object[]> results = marketingOrderRepo.findDistinctMonths();
        return results.stream()
                      .map(row -> new ViewDistinctMarketingOrder(
                          (Date) row[0], 
                          (Date) row[1], 
                          (Date) row[2]))
                      .collect(Collectors.toList());
    }
    
  //EXPORT RESUME
    public ByteArrayInputStream resumeMO(String month0, String month1, String month2) throws IOException {
    	List<MarketingOrder> marketingOrder = marketingOrderRepo.findMoAllTypeByMonth(month0, month1, month2);
    	List<HeaderMarketingOrder> headerMarketingOrder = headerMarketingOrderRepo.findByTwoMoId(marketingOrder.get(0).getMoId(), marketingOrder.get(1).getMoId());
    	List<DetailMarketingOrder> detailMarketingOrder = detailMarketingOrderRepo.findByTwoMoId(marketingOrder.get(0).getMoId(), marketingOrder.get(1).getMoId());
    	
    	String bulan0 = month0;
    	String bulan1 = month1;
    	String bulan2 = month2;
    	
    	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
    	try {
            Date date0 = dateFormat.parse(month0);
            Date date1 = dateFormat.parse(month1);
            Date date2 = dateFormat.parse(month2);

            // Format ulang menjadi 3 huruf pertama bulan
            SimpleDateFormat outputFormat0 = new SimpleDateFormat("MMM", Locale.ENGLISH);
            String m0 = outputFormat0.format(date0).toUpperCase();
            SimpleDateFormat outputFormat1 = new SimpleDateFormat("MMM", Locale.ENGLISH);
            String m1 = outputFormat1.format(date1).toUpperCase();
            SimpleDateFormat outputFormat2 = new SimpleDateFormat("MMM", Locale.ENGLISH);
            String m2 = outputFormat2.format(date2).toUpperCase();

            month0 = m0;
            month1 = m1;
            month2 = m2;
        } catch (Exception e) {
            e.printStackTrace();
        }
    	
    	String[] header = {
	            "Kategori", "Item", "Deskripsi", "Stock Awal",
	            "SF " + month0, "MO " + month0, "SF " + month1 , "MO " + month1, "SF " + month2, "MO " + month2
	        };
    	
        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("Sheet 1");
            
            // Set column width
            sheet.setColumnWidth(0, 4000);
            sheet.setColumnWidth(1, 5000);
            sheet.setColumnWidth(2, 11000);
            sheet.setColumnWidth(3, 3000);
            sheet.setColumnWidth(4, 3000);
            sheet.setColumnWidth(5, 3000);
            sheet.setColumnWidth(6, 3000);
            sheet.setColumnWidth(7, 3000);
            sheet.setColumnWidth(8, 3000);
            sheet.setColumnWidth(9, 3000);

            // Font
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
            
            Font calibriBold16 = workbook.createFont();
            calibriBold16.setFontName("Calibri");
            calibriBold16.setFontHeightInPoints((short) 16);
            calibriBold16.setBold(true);
            // End Font
            
            // Background Color
            XSSFColor lightBlueGray = new XSSFColor(new java.awt.Color(220, 230, 241), new DefaultIndexedColorMap());
            XSSFColor lightGray = new XSSFColor(new java.awt.Color(217, 217, 217), new DefaultIndexedColorMap());
            XSSFColor lightBlue = new XSSFColor(new java.awt.Color(155, 194, 230), new DefaultIndexedColorMap());
            XSSFColor lightRed = new XSSFColor(new java.awt.Color(255, 176, 132), new DefaultIndexedColorMap());
            XSSFColor gold = new XSSFColor(new java.awt.Color(255, 242, 204), new DefaultIndexedColorMap());
            XSSFColor yellow = new XSSFColor(new java.awt.Color(255, 255, 0), new DefaultIndexedColorMap());
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
            
            CellStyle lighBlueBorderStyle = workbook.createCellStyle();
            lighBlueBorderStyle.setBorderTop(BorderStyle.THIN);
            lighBlueBorderStyle.setBorderBottom(BorderStyle.THIN);;
            lighBlueBorderStyle.setTopBorderColor(IndexedColors.LIGHT_BLUE.getIndex());
            lighBlueBorderStyle.setBottomBorderColor(IndexedColors.LIGHT_BLUE.getIndex());;
            // End border cell style

            // Style untuk cell
            DataFormat dataFormat = workbook.createDataFormat();
            
            CellStyle calibri12Border = workbook.createCellStyle();
            calibri12Border.cloneStyleFrom(borderStyle);
            calibri12Border.setFont(calibri12);
            calibri12Border.setAlignment(HorizontalAlignment.LEFT);
            calibri12Border.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri12BorderRight = workbook.createCellStyle();
            calibri12BorderRight.cloneStyleFrom(borderStyle);
            calibri12BorderRight.setFont(calibri12);
            calibri12BorderRight.setAlignment(HorizontalAlignment.RIGHT);
            calibri12BorderRight.setVerticalAlignment(VerticalAlignment.CENTER);
            
            CellStyle calibri12BorderNum = workbook.createCellStyle();
            calibri12BorderNum.cloneStyleFrom(borderStyle);
            calibri12BorderNum.setFont(calibri12);
            calibri12BorderNum.setAlignment(HorizontalAlignment.RIGHT);
            calibri12BorderNum.setVerticalAlignment(VerticalAlignment.CENTER);
            calibri12BorderNum.setDataFormat(dataFormat.getFormat("#,##0"));
            
            CellStyle calibri12BorderNumYellow = workbook.createCellStyle();
            calibri12BorderNumYellow.cloneStyleFrom(borderStyle);
            calibri12BorderNumYellow.setFont(calibri12);
            calibri12BorderNumYellow.setAlignment(HorizontalAlignment.RIGHT);
            calibri12BorderNumYellow.setVerticalAlignment(VerticalAlignment.CENTER);
            calibri12BorderNumYellow.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibri12BorderNumYellow).setFillForegroundColor(yellow);
            calibri12BorderNumYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12BorderLightGray = workbook.createCellStyle();
            calibriBold12BorderLightGray.cloneStyleFrom(borderStyle);
            calibriBold12BorderLightGray.setFont(calibriBold12);
            calibriBold12BorderLightGray.setAlignment(HorizontalAlignment.CENTER);
            calibriBold12BorderLightGray.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12BorderLightGray).setFillForegroundColor(lightGray);
            calibriBold12BorderLightGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12BorderLightOrange = workbook.createCellStyle();
            calibriBold12BorderLightOrange.cloneStyleFrom(borderStyle);
            calibriBold12BorderLightOrange.setFont(calibriBold12);
            calibriBold12BorderLightOrange.setAlignment(HorizontalAlignment.LEFT);
            calibriBold12BorderLightOrange.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12BorderLightOrange).setFillForegroundColor(lightOrange);
            calibriBold12BorderLightOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibri12BoldBorderNumLightOrange = workbook.createCellStyle();
            calibri12BoldBorderNumLightOrange.cloneStyleFrom(borderStyle);
            calibri12BoldBorderNumLightOrange.setFont(calibriBold12);
            calibri12BoldBorderNumLightOrange.setAlignment(HorizontalAlignment.RIGHT);
            calibri12BoldBorderNumLightOrange.setVerticalAlignment(VerticalAlignment.CENTER);
            calibri12BoldBorderNumLightOrange.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibri12BoldBorderNumLightOrange).setFillForegroundColor(lightOrange);
            calibri12BoldBorderNumLightOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightRedLeft = workbook.createCellStyle();
            calibriBold12LightRedLeft.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightRedLeft.setFont(calibriBold12);
            calibriBold12LightRedLeft.setAlignment(HorizontalAlignment.LEFT);
            calibriBold12LightRedLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12LightRedLeft).setFillForegroundColor(lightRed);
            calibriBold12LightRedLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightRedRight = workbook.createCellStyle();
            calibriBold12LightRedRight.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightRedRight.setFont(calibriBold12);
            calibriBold12LightRedRight.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12LightRedRight.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12LightRedRight).setFillForegroundColor(lightRed);
            calibriBold12LightRedRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightRedNum = workbook.createCellStyle();
            calibriBold12LightRedNum.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightRedNum.setFont(calibriBold12);
            calibriBold12LightRedNum.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12LightRedNum.setVerticalAlignment(VerticalAlignment.CENTER);
            calibriBold12LightRedNum.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibriBold12LightRedNum).setFillForegroundColor(lightRed);
            calibriBold12LightRedNum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightBlueGrayLeft = workbook.createCellStyle();
            calibriBold12LightBlueGrayLeft.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightBlueGrayLeft.setFont(calibriBold12);
            calibriBold12LightBlueGrayLeft.setAlignment(HorizontalAlignment.LEFT);
            calibriBold12LightBlueGrayLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12LightBlueGrayLeft).setFillForegroundColor(lightBlueGray);
            calibriBold12LightBlueGrayLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightBlueGrayRight = workbook.createCellStyle();
            calibriBold12LightBlueGrayRight.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightBlueGrayRight.setFont(calibriBold12);
            calibriBold12LightBlueGrayRight.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12LightBlueGrayRight.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12LightBlueGrayRight).setFillForegroundColor(lightBlueGray);
            calibriBold12LightBlueGrayRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightBlueGrayNum = workbook.createCellStyle();
            calibriBold12LightBlueGrayNum.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightBlueGrayNum.setFont(calibriBold12);
            calibriBold12LightBlueGrayNum.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12LightBlueGrayNum.setVerticalAlignment(VerticalAlignment.CENTER);
            calibriBold12LightBlueGrayNum.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibriBold12LightBlueGrayNum).setFillForegroundColor(lightBlueGray);
            calibriBold12LightBlueGrayNum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12YellowLeft = workbook.createCellStyle();
            calibriBold12YellowLeft.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12YellowLeft.setFont(calibriBold12);
            calibriBold12YellowLeft.setAlignment(HorizontalAlignment.LEFT);
            calibriBold12YellowLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12YellowLeft).setFillForegroundColor(yellow);
            calibriBold12YellowLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12YellowRight = workbook.createCellStyle();
            calibriBold12YellowRight.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12YellowRight.setFont(calibriBold12);
            calibriBold12YellowRight.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12YellowRight.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12YellowRight).setFillForegroundColor(yellow);
            calibriBold12YellowRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12YellowNum = workbook.createCellStyle();
            calibriBold12YellowNum.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12YellowNum.setFont(calibriBold12);
            calibriBold12YellowNum.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12YellowNum.setVerticalAlignment(VerticalAlignment.CENTER);
            calibriBold12YellowNum.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibriBold12YellowNum).setFillForegroundColor(yellow);
            calibriBold12YellowNum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12YellowPercentage = workbook.createCellStyle();
            calibriBold12YellowPercentage.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12YellowPercentage.setFont(calibriBold12);
            calibriBold12YellowPercentage.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12YellowPercentage.setVerticalAlignment(VerticalAlignment.CENTER);
            calibriBold12YellowPercentage.setDataFormat(dataFormat.getFormat("#,##0.00\"%\""));
            ((XSSFCellStyle) calibriBold12YellowPercentage).setFillForegroundColor(yellow);
            calibriBold12YellowPercentage.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightBlueLeft = workbook.createCellStyle();
            calibriBold12LightBlueLeft.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightBlueLeft.setFont(calibriBold12);
            calibriBold12LightBlueLeft.setAlignment(HorizontalAlignment.LEFT);
            calibriBold12LightBlueLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12LightBlueLeft).setFillForegroundColor(lightBlue);
            calibriBold12LightBlueLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightBlueRight = workbook.createCellStyle();
            calibriBold12LightBlueRight.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightBlueRight.setFont(calibriBold12);
            calibriBold12LightBlueRight.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12LightBlueRight.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12LightBlueRight).setFillForegroundColor(lightBlue);
            calibriBold12LightBlueRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12LightBlueNum = workbook.createCellStyle();
            calibriBold12LightBlueNum.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12LightBlueNum.setFont(calibriBold12);
            calibriBold12LightBlueNum.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12LightBlueNum.setVerticalAlignment(VerticalAlignment.CENTER);
            calibriBold12LightBlueNum.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibriBold12LightBlueNum).setFillForegroundColor(lightBlue);
            calibriBold12LightBlueNum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12GoldLeft = workbook.createCellStyle();
            calibriBold12GoldLeft.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12GoldLeft.setFont(calibriBold12);
            calibriBold12GoldLeft.setAlignment(HorizontalAlignment.LEFT);
            calibriBold12GoldLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12GoldLeft).setFillForegroundColor(gold);
            calibriBold12GoldLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12GoldRight = workbook.createCellStyle();
            calibriBold12GoldRight.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12GoldRight.setFont(calibriBold12);
            calibriBold12GoldRight.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12GoldRight.setVerticalAlignment(VerticalAlignment.CENTER);
            ((XSSFCellStyle) calibriBold12GoldRight).setFillForegroundColor(gold);
            calibriBold12GoldRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibriBold12GoldNum = workbook.createCellStyle();
            calibriBold12GoldNum.cloneStyleFrom(lighBlueBorderStyle);
            calibriBold12GoldNum.setFont(calibriBold12);
            calibriBold12GoldNum.setAlignment(HorizontalAlignment.RIGHT);
            calibriBold12GoldNum.setVerticalAlignment(VerticalAlignment.CENTER);
            calibriBold12GoldNum.setDataFormat(dataFormat.getFormat("#,##0"));
            ((XSSFCellStyle) calibriBold12GoldNum).setFillForegroundColor(gold);
            calibriBold12GoldNum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            CellStyle calibri16Bold = workbook.createCellStyle();
            calibri16Bold.setFont(calibriBold16);
            calibri16Bold.setAlignment(HorizontalAlignment.LEFT);
            calibri16Bold.setVerticalAlignment(VerticalAlignment.CENTER);
            //end style

            // Header
            int row = 1;
            Row headerTableRow = sheet.createRow(row);
            headerTableRow.setHeight((short) 500);
            Cell headerTableCell;
            for (int i=0;i<header.length;i++) {
            	headerTableCell = headerTableRow.createCell(i);
            	headerTableCell.setCellValue(header[i]);
            	headerTableCell.setCellStyle(calibriBold12BorderLightGray);
            }
            
            // Daftar urutan kategori utama dan subkategori
            Map<String, List<String>> categoryOrder = new LinkedHashMap<>();
            categoryOrder.put("TUBE1", Collections.singletonList("FED TB NR"));
            categoryOrder.put("OEM", Arrays.asList("OEM TT", "OEM TL"));
            categoryOrder.put("HGP", Arrays.asList("HGP TT", "HGP TL"));
            categoryOrder.put("TUBE2", Collections.singletonList("FDR TB NR"));
            categoryOrder.put("FDR TT", Collections.singletonList("FDR TR TT"));
            categoryOrder.put("FDR TL", Collections.singletonList("FDR TR TL"));

            // Simpan data berdasarkan kategori utama
            Map<String, List<DetailMarketingOrder>> groupedData = new LinkedHashMap<>();

            // Kelompokkan data
            for (DetailMarketingOrder order : detailMarketingOrder) {
                String mainCategory = categoryOrder.keySet().stream()
                    .filter(key -> categoryOrder.get(key).contains(order.getCategory()))
                    .findFirst()
                    .orElse(order.getCategory()); // Jika kategori tidak ditemukan
                groupedData.computeIfAbsent(mainCategory, k -> new ArrayList<>()).add(order);
            }

            // Iterasi dan buat baris di Excel
            int currentRow = 2; // Baris pertama setelah header
            for (String mainCategory : categoryOrder.keySet()) {
                if (!groupedData.containsKey(mainCategory)) continue; // Lewati kategori tanpa data
                
                List<DetailMarketingOrder> orders = groupedData.get(mainCategory);
                Map<String, Double> totals = new HashMap<>(); // Menyimpan total untuk kategori utama
                
                for (String subCategory : categoryOrder.get(mainCategory)) {
                    for (DetailMarketingOrder order : orders) {
                        if (!order.getCategory().equals(subCategory)) continue;

                        Row dataRow = sheet.createRow(currentRow++);
                        Cell dataCell;

                        // Isi data sesuai urutan kolom
                        dataCell = dataRow.createCell(0);
                        dataCell.setCellValue(order.getCategory());
                        dataCell.setCellStyle(calibri12Border);

                        dataCell = dataRow.createCell(1);
                        dataCell.setCellValue(order.getPartNumber() != null ? order.getPartNumber().toString() : "");
                        dataCell.setCellStyle(calibri12Border);

                        dataCell = dataRow.createCell(2);
                        dataCell.setCellValue(order.getDescription() != null ? order.getDescription() : "");
                        dataCell.setCellStyle(calibri12Border);

                        double initialStock = order.getInitialStock() != null ? order.getInitialStock().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(3);
                        dataCell.setCellValue(initialStock);
                        dataCell.setCellStyle(calibri12BorderRight);

                        // Tambahkan ke total
                        totals.merge("InitialStock", initialStock, Double::sum);

                        // Lanjutkan untuk kolom lainnya
                        double sfMonth0 = order.getSfMonth0() != null ? order.getSfMonth0().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(4);
                        dataCell.setCellValue(sfMonth0);
                        dataCell.setCellStyle(calibri12BorderNumYellow);
                        totals.merge("SfMonth0", sfMonth0, Double::sum);

                        double moMonth0 = order.getMoMonth0() != null ? order.getMoMonth0().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(5);
                        dataCell.setCellValue(moMonth0);
                        dataCell.setCellStyle(calibri12BorderNum);
                        totals.merge("MoMonth0", moMonth0, Double::sum);

                        double sfMonth1 = order.getSfMonth1() != null ? order.getSfMonth1().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(6);
                        dataCell.setCellValue(sfMonth1);
                        dataCell.setCellStyle(calibri12BorderNumYellow);
                        totals.merge("SfMonth1", sfMonth1, Double::sum);

                        double moMonth1 = order.getMoMonth1() != null ? order.getMoMonth1().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(7);
                        dataCell.setCellValue(moMonth1);
                        dataCell.setCellStyle(calibri12BorderNum);
                        totals.merge("MoMonth1", moMonth1, Double::sum);

                        double sfMonth2 = order.getSfMonth2() != null ? order.getSfMonth2().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(8);
                        dataCell.setCellValue(sfMonth2);
                        dataCell.setCellStyle(calibri12BorderNumYellow);
                        totals.merge("SfMonth2", sfMonth2, Double::sum);

                        double moMonth2 = order.getMoMonth2() != null ? order.getMoMonth2().doubleValue() : 0.0;
                        dataCell = dataRow.createCell(9);
                        dataCell.setCellValue(moMonth2);
                        dataCell.setCellStyle(calibri12BorderNum);
                        totals.merge("MoMonth2", moMonth2, Double::sum);
                    }
                }

                // Tambahkan baris jumlah setelah subkategori selesai
                Row totalRow = sheet.createRow(currentRow++);
                Cell totalCell = totalRow.createCell(0);
                totalCell.setCellValue("");
                totalCell.setCellStyle(calibriBold12BorderLightOrange);
                
                totalCell = totalRow.createCell(1);
                totalCell.setCellValue("TOTAL " + mainCategory);
                totalCell.setCellStyle(calibriBold12BorderLightOrange);

                totalCell = totalRow.createCell(2);
                totalCell.setCellValue("");
                totalCell.setCellStyle(calibriBold12BorderLightOrange);
                
                totalCell = totalRow.createCell(3);
                totalCell.setCellValue(totals.getOrDefault("InitialStock", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);

                totalCell = totalRow.createCell(4);
                totalCell.setCellValue(totals.getOrDefault("SfMonth0", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);

                totalCell = totalRow.createCell(5);
                totalCell.setCellValue(totals.getOrDefault("MoMonth0", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);

                totalCell = totalRow.createCell(6);
                totalCell.setCellValue(totals.getOrDefault("SfMonth1", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);

                totalCell = totalRow.createCell(7);
                totalCell.setCellValue(totals.getOrDefault("MoMonth1", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);

                totalCell = totalRow.createCell(8);
                totalCell.setCellValue(totals.getOrDefault("SfMonth2", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);

                totalCell = totalRow.createCell(9);
                totalCell.setCellValue(totals.getOrDefault("MoMonth2", 0.0));
                totalCell.setCellStyle(calibri12BoldBorderNumLightOrange);
            }
            
            String[] summaryHeader = {
		            "BRAND", "", "", "", "SF " + month0, "MO " + month0, "SF " + month1, "MO " + month1, "SF " + month2, "MO " + month2
		        };
            
            CellStyle[] summaryStyleHeader = {
		            calibriBold12LightRedLeft, calibriBold12LightRedLeft, calibriBold12LightRedLeft, 
		            calibriBold12LightRedLeft, calibriBold12LightRedRight, calibriBold12LightRedRight,
		            calibriBold12LightRedRight, calibriBold12LightRedRight, calibriBold12LightRedRight, 
		            calibriBold12LightRedRight
		        };
            
            String[] tireSummary = {
		            "OEM TT", "OEM TL", "TOTAL OEM", "HGP TT", "HGP TL", "TOTAL HGP", "FDR TR TT", "FDR TR TL", 
		            "TOTAL FEDERAL", "TOTAL FDR", "TOTAL TIRE", "TIRE TT", "TIRE TL", "% TT", "% TL", "OEM TL", "HGP TL", 
		            "FEDERAL TL", "FDR TL", "FED", "FDR", "MO VS SF FED", "MO VS SF FDR"
		        };
            
            CellStyle[] tireSummaryStyle = {
		            calibriBold12LightBlueGrayLeft, calibriBold12LightBlueGrayLeft, calibriBold12YellowLeft, 
		            calibriBold12LightBlueGrayLeft, calibriBold12LightBlueGrayLeft, calibriBold12YellowLeft,
		            calibriBold12LightBlueGrayLeft, calibriBold12LightBlueGrayLeft, calibriBold12LightBlueLeft, 
		            calibriBold12LightBlueLeft, calibriBold12LightBlueLeft, calibriBold12LightBlueLeft, 
		            calibriBold12LightBlueLeft, calibriBold12YellowLeft, calibriBold12YellowLeft, 
		            calibriBold12YellowLeft, calibriBold12YellowLeft, calibriBold12YellowLeft, 
		            calibriBold12YellowLeft, calibriBold12LightBlueLeft, calibriBold12LightBlueLeft, 
		            calibriBold12GoldLeft, calibriBold12GoldLeft
		        };
            
            CellStyle[] tireNumSummaryStyle = {
		            calibriBold12LightBlueGrayNum, calibriBold12LightBlueGrayNum, calibriBold12YellowNum, 
		            calibriBold12LightBlueGrayNum, calibriBold12LightBlueGrayNum, calibriBold12YellowNum,
		            calibriBold12LightBlueGrayNum, calibriBold12LightBlueGrayNum, calibriBold12LightBlueNum, 
		            calibriBold12LightBlueNum, calibriBold12LightBlueNum, calibriBold12LightBlueNum, 
		            calibriBold12LightBlueNum, calibriBold12YellowNum, calibriBold12YellowNum, 
		            calibriBold12YellowNum, calibriBold12YellowNum, calibriBold12YellowNum, 
		            calibriBold12YellowNum, calibriBold12LightBlueNum, calibriBold12LightBlueNum, 
		            calibriBold12GoldNum, calibriBold12GoldNum
		        };
            
            currentRow += 4;
            Row tireRow = sheet.createRow(currentRow);
            Cell tireCell = tireRow.createCell(0);
            tireCell.setCellValue("TIRE");
            tireCell.setCellStyle(calibri16Bold);
            
            currentRow += 1;
            Row summaryRow = sheet.createRow(currentRow);
            for(int i=0;i<summaryHeader.length;i++) {
            	Cell summaryCell = summaryRow.createCell(i);
                summaryCell.setCellValue(summaryHeader[i]);
                summaryCell.setCellStyle(summaryStyleHeader[i]);
            }
            
            currentRow += 1;
            for(int i=0;i<tireSummary.length;i++) {
            	summaryRow = sheet.createRow(currentRow);
                Cell summaryCell = summaryRow.createCell(0);
                summaryCell.setCellValue(tireSummary[i]);
                summaryCell.setCellStyle(tireSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(1);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tireSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(2);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tireSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(3);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tireSummaryStyle[i]);
                
                if (i == 2) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 5) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 8) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 9) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 10) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 11) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 12) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TL", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TL", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TL", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TL", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TL", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TL", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 13) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 0)
                    + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0) + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) 
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1)) /(sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                } else if (i == 14) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TL", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TL", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TL", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TL", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TL", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, "OEM TL", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                } else if (i == 15) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "sf", 0)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) 
                    + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "mo", 0)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) 
                    + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "sf", 1)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) 
                    + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "mo", 1)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) 
                    + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "sf", 2)) / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) 
                    + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "mo", 2)) / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) 
                    + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                } else if (i == 16) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[4], "sf", 0)) / (sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) 
                    + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[4], "mo", 0)) / (sumTotal(detailMarketingOrder, "HGP TT", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[4], "sf", 1)) / (sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) 
                    + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[4], "mo", 1)) / (sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) 
                    + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[4], "sf", 2)) / (sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) 
                    + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[4], "mo", 2)) / (sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) 
                    + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                } else if (i == 17) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "sf", 0) + sumTotal(detailMarketingOrder, tireSummary[4], "sf", 0)) 
                    / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "mo", 0) + sumTotal(detailMarketingOrder, tireSummary[4], "mo", 0)) 
                    / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "sf", 1) + sumTotal(detailMarketingOrder, tireSummary[4], "sf", 1)) 
                    / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "mo", 1) + sumTotal(detailMarketingOrder, tireSummary[4], "mo", 1)) 
                    / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "sf", 2) + sumTotal(detailMarketingOrder, tireSummary[4], "sf", 2))
                    / (sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[1], "mo", 2) + sumTotal(detailMarketingOrder, tireSummary[4], "mo", 2)) 
                    / (sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                } else if (i == 18) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[7], "sf", 0)) / (sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) 
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[7], "mo", 0)) / (sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) 
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[7], "sf", 1)) / (sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1))) * 100);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[7], "mo", 1)) / (sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) 
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue((sumTotal(detailMarketingOrder, tireSummary[7], "sf", 2)) / (sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2)
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2)) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(((sumTotal(detailMarketingOrder, tireSummary[7], "mo", 2)) / (sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) 
                    + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2))) * 100);
                    summaryCell.setCellStyle(calibriBold12YellowPercentage);
                } else if (i == 19) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 20) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 21) {
                	DecimalFormatSymbols symbols = new DecimalFormatSymbols();
                    symbols.setGroupingSeparator('.');
                	DecimalFormat formatter = new DecimalFormat("#,###", symbols);
                	Double vs0 = (sumTotal(detailMarketingOrder, "OEM TT", "sf", 0) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 0) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 0)) - (sumTotal(detailMarketingOrder, "OEM TT", "mo", 0) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 0)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 0) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 0));
                	String resultVs0 = "(" + formatter.format(vs0) + ")";
                	
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(resultVs0);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue("");
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    Double vs1 = (sumTotal(detailMarketingOrder, "OEM TT", "sf", 1) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 1) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 1)) - (sumTotal(detailMarketingOrder, "OEM TT", "mo", 1) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 1)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 1) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 1));
                	String resultVs1 = "(" + formatter.format(vs1) + ")";
                	
                	summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(resultVs1);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue("");
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    Double vs2 = (sumTotal(detailMarketingOrder, "OEM TT", "sf", 2) + sumTotal(detailMarketingOrder, "OEM TL", "sf", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "sf", 2) + sumTotal(detailMarketingOrder, "HGP TL", "sf", 2)) - (sumTotal(detailMarketingOrder, "OEM TT", "mo", 2) + sumTotal(detailMarketingOrder, "OEM TL", "mo", 2)
                    + sumTotal(detailMarketingOrder, "HGP TT", "mo", 2) + sumTotal(detailMarketingOrder, "HGP TL", "mo", 2));
                	String resultVs2 = "(" + formatter.format(vs2) + ")";
                	
                	summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(resultVs2);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue("");
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else if (i == 22) {
                	DecimalFormatSymbols symbols = new DecimalFormatSymbols();
                    symbols.setGroupingSeparator('.');
                	DecimalFormat formatter = new DecimalFormat("#,###", symbols);
                	Double vs0 = (sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 0)) 
                	- (sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 0) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 0));
                	String resultVs0 = "(" + formatter.format(vs0) + ")";
                	
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(resultVs0);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue("");
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    Double vs1 = (sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 1))
                    - (sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 1) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 1));
                	String resultVs1 = "(" + formatter.format(vs1) + ")";
                	
                	summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(resultVs1);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue("");
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    Double vs2 = (sumTotal(detailMarketingOrder, "FDR TR TT", "sf", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "sf", 2)) 
                    - (sumTotal(detailMarketingOrder, "FDR TR TT", "mo", 2) + sumTotal(detailMarketingOrder, "FDR TR TL", "mo", 2));
                	String resultVs2 = "(" + formatter.format(vs2) + ")";
                	
                	summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(resultVs2);
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue("");
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                } else {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "sf", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "mo", 0));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "sf", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "mo", 1));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "sf", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "mo", 2));
                    summaryCell.setCellStyle(tireNumSummaryStyle[i]);
                }
                currentRow++;
            }
            
            String[] tubeSummary = {
		            "FDR TB NR", "TOTAL FDR", "FED TB NR", "OEM TT", "TOTAL FEDERAL", "TOTAL TUBE"
		        };
            
            CellStyle[] tubeSummaryStyle = {
		            calibriBold12LightBlueGrayLeft, calibriBold12YellowLeft, calibriBold12LightBlueGrayLeft, 
		            calibriBold12LightBlueGrayLeft, calibriBold12YellowLeft, calibriBold12LightBlueLeft
		        };
            
            CellStyle[] tubeNumSummaryStyle = {
		            calibriBold12LightBlueGrayNum, calibriBold12YellowNum, calibriBold12LightBlueGrayNum, 
		            calibriBold12LightBlueGrayNum, calibriBold12YellowNum, calibriBold12LightBlueNum
		        };
            
            currentRow += 2;
            Row tubeRow = sheet.createRow(currentRow);
            Cell tubeCell = tubeRow.createCell(0);
            tubeCell.setCellValue("TUBE");
            tubeCell.setCellStyle(calibri16Bold);
            
            currentRow += 1;
            summaryRow = sheet.createRow(currentRow);
            for(int i=0;i<summaryHeader.length;i++) {
            	Cell summaryCell = summaryRow.createCell(i);
                summaryCell.setCellValue(summaryHeader[i]);
                summaryCell.setCellStyle(summaryStyleHeader[i]);
            }
            
            currentRow += 1;
            summaryRow = sheet.createRow(currentRow);
            for(int i=0;i<tubeSummary.length;i++) {
            	summaryRow = sheet.createRow(currentRow);
            	Cell summaryCell = summaryRow.createCell(0);
                summaryCell.setCellValue(tubeSummary[i]);
                summaryCell.setCellStyle(tubeSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(1);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tubeSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(2);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tubeSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(3);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tubeSummaryStyle[i]);
                
                summaryCell = summaryRow.createCell(4);
                summaryCell.setCellValue("");
                summaryCell.setCellStyle(tubeSummaryStyle[i]);
                
                if (i == 1) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "sf", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "mo", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "sf", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "mo", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "sf", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "mo", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                } else if (i == 4) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[2], "sf", 0) + sumTotal(detailMarketingOrder, tireSummary[3], "sf", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[2], "mo", 0) + sumTotal(detailMarketingOrder, tireSummary[3], "mo", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[2], "sf", 1) + sumTotal(detailMarketingOrder, tireSummary[3], "sf", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[2], "mo", 1) + sumTotal(detailMarketingOrder, tireSummary[3], "mo", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[2], "sf", 2) + sumTotal(detailMarketingOrder, tireSummary[3], "sf", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[2], "mo", 2) + sumTotal(detailMarketingOrder, tireSummary[3], "mo", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                } else if (i == 5) {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "sf", 0) + sumTotal(detailMarketingOrder, tireSummary[2], "sf", 0) + sumTotal(detailMarketingOrder, tireSummary[3], "sf", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "mo", 0) + sumTotal(detailMarketingOrder, tireSummary[2], "mo", 0) + sumTotal(detailMarketingOrder, tireSummary[3], "mo", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "sf", 1) + sumTotal(detailMarketingOrder, tireSummary[2], "sf", 1) + sumTotal(detailMarketingOrder, tireSummary[3], "sf", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "mo", 1) + sumTotal(detailMarketingOrder, tireSummary[2], "mo", 1) + sumTotal(detailMarketingOrder, tireSummary[3], "mo", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "sf", 2) + sumTotal(detailMarketingOrder, tireSummary[2], "sf", 2) + sumTotal(detailMarketingOrder, tireSummary[3], "sf", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[0], "mo", 2) + sumTotal(detailMarketingOrder, tireSummary[2], "mo", 2) + sumTotal(detailMarketingOrder, tireSummary[3], "mo", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                } else {
                	summaryCell = summaryRow.createCell(4);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "sf", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(5);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "mo", 0));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(6);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "sf", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(7);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "mo", 1));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(8);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "sf", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                    
                    summaryCell = summaryRow.createCell(9);
                    summaryCell.setCellValue(sumTotal(detailMarketingOrder, tireSummary[i], "mo", 2));
                    summaryCell.setCellStyle(tubeNumSummaryStyle[i]);
                }
                
                currentRow++;
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
    
    private Double sumTotal(List<DetailMarketingOrder> marketingOrder, String category, String what, int month) {
        Double total = 0.0;

        for (DetailMarketingOrder order : marketingOrder) {
            if (order.getCategory() != null && order.getCategory().equals(category)) {
                if ("sf".equals(what)) {
                    switch (month) {
                        case 0:
                            total += order.getSfMonth0() != null ? order.getSfMonth0().doubleValue() : 0.0;
                            break;
                        case 1:
                            total += order.getSfMonth1() != null ? order.getSfMonth1().doubleValue() : 0.0;
                            break;
                        case 2:
                            total += order.getSfMonth2() != null ? order.getSfMonth2().doubleValue() : 0.0;
                            break;
                    }
                } else if ("mo".equals(what)) {
                    switch (month) {
                        case 0:
                            total += order.getMoMonth0() != null ? order.getMoMonth0().doubleValue() : 0.0;
                            break;
                        case 1:
                            total += order.getMoMonth1() != null ? order.getMoMonth1().doubleValue() : 0.0;
                            break;
                        case 2:
                            total += order.getMoMonth2() != null ? order.getMoMonth2().doubleValue() : 0.0;
                            break;
                    }
                }
            }
        }

        return total;
    }
    
    //Save Ar Defect Reject
    public int saveArDefectReject(SaveFinalMarketingOrder mo) {
    	String moIdFed = "";
    	moIdFed = getLastIdMo();
    	
        List<DetailMarketingOrder> detailTempFed = new ArrayList<>();
        List<ItemCuring> curingList = itemCuringRepo.findAll();
        List<Product> prodList = productRepo.findAll();
        List<ProductType> prodTypeList = productTypeRepo.findAll();
        
        BigDecimal fedabM0 = BigDecimal.ZERO;
        BigDecimal fedabM1 = BigDecimal.ZERO;
        BigDecimal fedabM2 = BigDecimal.ZERO;
        BigDecimal fedtlM0 = BigDecimal.ZERO;
        BigDecimal fedtlM1 = BigDecimal.ZERO;
        BigDecimal fedtlM2 = BigDecimal.ZERO;
        BigDecimal fedttM0 = BigDecimal.ZERO;
        BigDecimal fedttM1 = BigDecimal.ZERO;
        BigDecimal fedttM2 = BigDecimal.ZERO;
        String itemCuringFed = " ";
    	
        // Start Save Detail Marketing Order FED
        
        MarketingOrder dataMarketingFed = mo.getMoFed();
        List<HeaderMarketingOrder> dataHeaderFed = mo.getHeaderMoFed();
    	List<DetailMarketingOrder> dataDetailFed = mo.getDetailMoFed();
    	for(DetailMarketingOrder data: dataDetailFed) {
    		DetailMarketingOrder detailMo = new DetailMarketingOrder();
    		BigDecimal detailId = getNewDetailMarketingOrderId();
    		
            detailMo.setDetailId(detailId);
            detailMo.setMoId(moIdFed);
            detailMo.setDescription(data.getDescription());
            detailMo.setCategory(data.getCategory());
            detailMo.setMachineType(data.getMachineType());
            detailMo.setPartNumber(data.getPartNumber());
            detailMo.setCapacity(data.getCapacity());
            detailMo.setQtyPerMould(data.getQtyPerMould());
            detailMo.setQtyPerRak(data.getQtyPerRak());
            detailMo.setMinOrder(data.getMinOrder());
            detailMo.setMaxCapMonth0(data.getMaxCapMonth0());
            detailMo.setMaxCapMonth1(data.getMaxCapMonth1());
            detailMo.setMaxCapMonth2(data.getMaxCapMonth2());
            detailMo.setInitialStock(data.getInitialStock());
            detailMo.setSfMonth0(data.getSfMonth0());
            detailMo.setSfMonth1(data.getSfMonth1());
            detailMo.setSfMonth2(data.getSfMonth2());
            detailMo.setMoMonth0(data.getMoMonth0());
            detailMo.setMoMonth1(data.getMoMonth1());
            detailMo.setMoMonth2(data.getMoMonth2());
            detailMo.setLockStatusM0(data.getLockStatusM0());
            detailMo.setLockStatusM1(data.getLockStatusM1());
            detailMo.setLockStatusM2(data.getLockStatusM2());
            detailMo.setTotalAr(data.getTotalAr());
            detailMo.setAr(data.getAr());
            detailMo.setDefect(data.getDefect());
            detailMo.setReject(data.getReject());
            
            BigDecimal totalMO = detailMo.getMoMonth0();
            BigDecimal prodType = BigDecimal.ZERO;
            BigDecimal hk = BigDecimal.ZERO;
            BigDecimal ppd = BigDecimal.ZERO;
            
            for (Product pro : prodList) {
                if (pro.getPART_NUMBER().equals(detailMo.getPartNumber())) {
                    prodType = pro.getPRODUCT_TYPE_ID();
                    itemCuringFed = pro.getITEM_CURING();
                    break; 
                }
            }
            
            BigDecimal HKTT = dataHeaderFed.get(0).getTotalWdTt(); 
            BigDecimal HKTL = dataHeaderFed.get(0).getTotalWdTl(); 
            
        	for (ProductType prot : prodTypeList) {
                if (prot.getPRODUCT_TYPE_ID().equals(prodType)) {
                    if (prot.getPRODUCT_TYPE().equals("TT")) {
                        hk = HKTT;
                        fedttM0 = fedttM0.add(data.getMoMonth0());
                        fedttM1 = fedttM1.add(data.getMoMonth1());
                        fedttM2 = fedttM2.add(data.getMoMonth2());
                    } else if (prot.getPRODUCT_TYPE().equals("TL")) {
                        hk = HKTL;
                        fedtlM0 = fedtlM0.add(data.getMoMonth0());
                        fedtlM1 = fedtlM1.add(data.getMoMonth1());
                        fedtlM2 = fedtlM2.add(data.getMoMonth2());
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
            BigDecimal cav = ppd.divide(data.getCapacity(), RoundingMode.HALF_UP);
            detailMo.setCav(cav.compareTo(BigDecimal.ONE) < 0 ? BigDecimal.ONE : cav);
            
            // Hitung Airbag Machine
            for (ItemCuring cur : curingList) {
                if (cur.getITEM_CURING().compareTo(itemCuringFed) == 0) {
                    if (cur.getMACHINE_TYPE().compareTo("A/B") == 0) {
                        fedabM0 = fedabM0.add(data.getMoMonth0());
                        fedabM1 = fedabM1.add(data.getMoMonth1());
                        fedabM2 = fedabM2.add(data.getMoMonth2());
                    }
                }
            }
            
            detailTempFed.add(detailMarketingOrderRepo.save(detailMo));
    	}
    	
    	//End Save Detail Marketing Order FED
    	
    	// Save Header Marketing Order FED
    	
        for (HeaderMarketingOrder hmo : dataHeaderFed) {
        	BigDecimal newHeaderId = getNewHeaderMarketingOrderId();
        	hmo.setHeaderId(newHeaderId);
        	hmo.setMoId(moIdFed);
        	hmo.setStatus(BigDecimal.ONE);
        	
            if (hmo.getMonth().equals(dataMarketingFed.getMonth0())) {
                hmo.setAirbagMachine(fedabM0);
                hmo.setTl(fedtlM0);
                hmo.setTt(fedttM0);
                hmo.setTotalMo(fedtlM0.add(fedttM0));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(dataMarketingFed.getMonth1())) {
                hmo.setAirbagMachine(fedabM1);
                hmo.setTl(fedtlM1);
                hmo.setTt(fedttM1);
                hmo.setTotalMo(fedtlM1.add(fedttM1));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(dataMarketingFed.getMonth2())) {
                hmo.setAirbagMachine(fedabM2);
                hmo.setTl(fedtlM2);
                hmo.setTt(fedttM2);
                hmo.setTotalMo(fedtlM2.add(fedttM2));
                updatePercentages(hmo);
            }
            headerMarketingOrderRepo.save(hmo);
        }
        
        // End Save Header Marketing Order FED
        
        
        // Start Save Marketing Order FED
        
        if(dataMarketingFed != null) {
        	dataMarketingFed.setStatusFilled(BigDecimal.valueOf(3));
        	dataMarketingFed.setMoId(moIdFed);
		}
		
        if (dataMarketingFed.getRevisionPpc() == null) {
        	dataMarketingFed.setRevisionPpc(BigDecimal.ONE); // Jika null atau 0, set ke 1
        } else {
        	dataMarketingFed.setRevisionPpc(dataMarketingFed.getRevisionPpc().add(BigDecimal.ONE)); // Tambah 1 pada revisi
        }
        	        
        marketingOrderRepo.save(dataMarketingFed);
        
       // End Save Marketing Order FED
        
       
	    String moIdFdr = "";
    	moIdFdr = getLastIdMo();
    	
        List<DetailMarketingOrder> detailTempFdr = new ArrayList<>();
        
        BigDecimal fdrabM0 = BigDecimal.ZERO;
        BigDecimal fdrabM1 = BigDecimal.ZERO;
        BigDecimal fdrabM2 = BigDecimal.ZERO;
        BigDecimal fdrtlM0 = BigDecimal.ZERO;
        BigDecimal fdrtlM1 = BigDecimal.ZERO;
        BigDecimal fdrtlM2 = BigDecimal.ZERO;
        BigDecimal fdrttM0 = BigDecimal.ZERO;
        BigDecimal fdrttM1 = BigDecimal.ZERO;
        BigDecimal fdrttM2 = BigDecimal.ZERO;
        String itemCuringFdr = "";
    	
        // Start Save Detail Marketing Order FDR
        
        MarketingOrder dataMarketingFdr = mo.getMoFdr();
        List<HeaderMarketingOrder> dataHeaderFdr = mo.getHeaderMoFdr();
    	List<DetailMarketingOrder> dataDetailFdr = mo.getDetailMoFdr();
    	for(DetailMarketingOrder data: dataDetailFdr) {
    		DetailMarketingOrder detailMo = new DetailMarketingOrder();
    		BigDecimal detailId = getNewDetailMarketingOrderId();
    		
            detailMo.setDetailId(detailId);
            detailMo.setMoId(moIdFdr);
            detailMo.setDescription(data.getDescription());
            detailMo.setCategory(data.getCategory());
            detailMo.setMachineType(data.getMachineType());
            detailMo.setPartNumber(data.getPartNumber());
            detailMo.setCapacity(data.getCapacity());
            detailMo.setQtyPerMould(data.getQtyPerMould());
            detailMo.setQtyPerRak(data.getQtyPerRak());
            detailMo.setMinOrder(data.getMinOrder());
            detailMo.setMaxCapMonth0(data.getMaxCapMonth0());
            detailMo.setMaxCapMonth1(data.getMaxCapMonth1());
            detailMo.setMaxCapMonth2(data.getMaxCapMonth2());
            detailMo.setInitialStock(data.getInitialStock());
            detailMo.setSfMonth0(data.getSfMonth0());
            detailMo.setSfMonth1(data.getSfMonth1());
            detailMo.setSfMonth2(data.getSfMonth2());
            detailMo.setMoMonth0(data.getMoMonth0());
            detailMo.setMoMonth1(data.getMoMonth1());
            detailMo.setMoMonth2(data.getMoMonth2());
            detailMo.setLockStatusM0(data.getLockStatusM0());
            detailMo.setLockStatusM1(data.getLockStatusM1());
            detailMo.setLockStatusM2(data.getLockStatusM2());
            detailMo.setTotalAr(data.getTotalAr());
            detailMo.setAr(data.getAr());
            detailMo.setDefect(data.getDefect());
            detailMo.setReject(data.getReject());
            
            BigDecimal totalMO = detailMo.getMoMonth0();
            BigDecimal prodType = BigDecimal.ZERO;
            BigDecimal hk = BigDecimal.ZERO;
            BigDecimal ppd = BigDecimal.ZERO;
            
            for (Product pro : prodList) {
                if (pro.getPART_NUMBER().equals(detailMo.getPartNumber())) {
                    prodType = pro.getPRODUCT_TYPE_ID();
                    itemCuringFdr = pro.getITEM_CURING();
                    break; 
                }
            }
            
            BigDecimal HKTT = dataHeaderFdr.get(0).getTotalWdTt(); 
            BigDecimal HKTL = dataHeaderFdr.get(0).getTotalWdTl(); 
            
        	for (ProductType prot : prodTypeList) {
                if (prot.getPRODUCT_TYPE_ID().equals(prodType)) {
                    if (prot.getPRODUCT_TYPE().equals("TT")) {
                        hk = HKTT;
                        fdrttM0 = fdrttM0.add(data.getMoMonth0());
                        fdrttM1 = fdrttM1.add(data.getMoMonth1());
                        fdrttM2 = fdrttM2.add(data.getMoMonth2());
                    } else if (prot.getPRODUCT_TYPE().equals("TL")) {
                        hk = HKTL;
                        fdrtlM0 = fdrtlM0.add(data.getMoMonth0());
                        fdrtlM1 = fdrtlM1.add(data.getMoMonth1());
                        fdrtlM2 = fdrtlM2.add(data.getMoMonth2());
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
            BigDecimal cav = ppd.divide(data.getCapacity(), RoundingMode.HALF_UP);
            detailMo.setCav(cav.compareTo(BigDecimal.ONE) < 0 ? BigDecimal.ONE : cav);
            
            // Hitung Airbag Machine
            for (ItemCuring cur : curingList) {
                if (cur.getITEM_CURING().compareTo(itemCuringFdr) == 0) {
                    if (cur.getMACHINE_TYPE().compareTo("A/B") == 0) {
                        fdrabM0 = fdrabM0.add(data.getMoMonth0());
                        fdrabM1 = fdrabM1.add(data.getMoMonth1());
                        fdrabM2 = fdrabM2.add(data.getMoMonth2());
                    }
                }
            }
            
            detailTempFdr.add(detailMarketingOrderRepo.save(detailMo));
    	}
    	
    	//End Save Detail Marketing Order FDR
    	
    	// Save Header Marketing Order FDR
    	
        for (HeaderMarketingOrder hmo : dataHeaderFdr) {
        	BigDecimal newHeaderId = getNewHeaderMarketingOrderId();
        	hmo.setHeaderId(newHeaderId);
        	hmo.setMoId(moIdFdr);
        	hmo.setStatus(BigDecimal.ONE);
        	
            if (hmo.getMonth().equals(dataMarketingFdr.getMonth0())) {
                hmo.setAirbagMachine(fdrabM0);
                hmo.setTl(fdrtlM0);
                hmo.setTt(fdrttM0);
                hmo.setTotalMo(fdrtlM0.add(fdrttM0));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(dataMarketingFdr.getMonth1())) {
                hmo.setAirbagMachine(fdrabM1);
                hmo.setTl(fdrtlM1);
                hmo.setTt(fdrttM1);
                hmo.setTotalMo(fdrtlM1.add(fdrttM1));
                updatePercentages(hmo);
            } else if (hmo.getMonth().equals(dataMarketingFdr.getMonth2())) {
                hmo.setAirbagMachine(fdrabM2);
                hmo.setTl(fdrtlM2);
                hmo.setTt(fdrttM2);
                hmo.setTotalMo(fdrtlM2.add(fdrttM2));
                updatePercentages(hmo);
            }
            headerMarketingOrderRepo.save(hmo);
        }
        
        // End Save Header Marketing Order FDR
        
        
        // Start Save Marketing Order FDR
        
        if(dataMarketingFdr != null) {
        	dataMarketingFdr.setStatusFilled(BigDecimal.valueOf(3));
        	dataMarketingFdr.setMoId(moIdFdr);
		}
		
        if (dataMarketingFdr.getRevisionPpc() == null) {
        	dataMarketingFdr.setRevisionPpc(BigDecimal.ONE);// Jika null atau 0, set ke 1
        } else {
        	dataMarketingFdr.setRevisionPpc(dataMarketingFdr.getRevisionPpc().add(BigDecimal.ONE)); // Tambah 1 pada revisi
        }
        	        
        marketingOrderRepo.save(dataMarketingFdr);
        
       // End Save Marketing Order FDR
    	
    	return 1;
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
    
	//GET ALL MARKETING ORDER ONLY MONTH
    public List<Map<String, Object>> getAllMoOnlyMonth() {
        return marketingOrderRepo.findOnlyMonth();
    }
    
    
    //Get MO FED & FDR Marketing Order by Month
    public GetAllTypeMarketingOrder getAllTypeMarketingOrder(String dateMoMonth0, String dateMoMonth1, String dateMoMonth2) {
    	
    	GetAllTypeMarketingOrder result = new GetAllTypeMarketingOrder();
    	
    	//Search data MO FED FDR
    	String moIdFed = null;
    	String moIdFdr = null;
    	List<MarketingOrder> dataMo = findMoAllTypeByMonth(dateMoMonth0, dateMoMonth1, dateMoMonth2);
    	
    	//Loop cek berdasarkan type dan set marketing ordernya
    	for (MarketingOrder mo : dataMo) {
    	    if ("FDR".equals(mo.getType())) {
    	    	moIdFdr = mo.getMoId();
    	    	result.setMoFdr(mo);
    	    } else if ("FED".equals(mo.getType())) {
    	    	moIdFed = mo.getMoId();
    	    	result.setMoFed(mo);
    	    }
    	}
    	
    	//List Product Untuk Item Curing
    	List<Product> prodList = productRepo.findAll();
    	    	
    	
    	//Set Header & Detail FED 
        List<HeaderMarketingOrder> hmoFed = headerMarketingOrderRepo.findByMoId(moIdFed);
        List<DetailMarketingOrder> dmoFed = detailMarketingOrderRepo.findByMoId(moIdFed);
        List<ViewDetailMarketingOrder> detailResponsesFed = new ArrayList<>();
        for (DetailMarketingOrder detail : dmoFed) {
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

        	String itemCuring = null; 
        	for (Product product : prodList) {
        	    if (product.getPART_NUMBER().equals(detail.getPartNumber())) {
        	        itemCuring = product.getITEM_CURING(); 
        	        break;  
        	    }
        	}

        	detailResponse.setItemCuring(itemCuring);
            detailResponsesFed.add(detailResponse);
        }
        
        result.setHeaderMarketingOrderFed(hmoFed);
    	result.setDetailMarketingOrderFed(detailResponsesFed);
        
    	//Set header & Detail FDR
        List<HeaderMarketingOrder> hmoFdr = headerMarketingOrderRepo.findByMoId(moIdFdr);
        List<DetailMarketingOrder> dmoFdr = detailMarketingOrderRepo.findByMoId(moIdFdr);
        List<ViewDetailMarketingOrder> detailResponsesFdr = new ArrayList<>();
        for (DetailMarketingOrder detail : dmoFdr) {
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

        	String itemCuring = null; 
        	for (Product product : prodList) {
        	    if (product.getPART_NUMBER().equals(detail.getPartNumber())) {
        	        itemCuring = product.getITEM_CURING(); 
        	        break;  
        	    }
        	}

        	detailResponse.setItemCuring(itemCuring);
            detailResponsesFdr.add(detailResponse);
        }
    	
    	result.setHeaderMarketingOrderFdr(hmoFdr);
    	result.setDetailMarketingOrderFdr(detailResponsesFdr);
    	
    	return result;
    }
    
    public List<MarketingOrder> findMoAllTypeByMonth(String dateMoMonth0, String dateMoMonth1, String dateMoMonth2){
    	List<MarketingOrder> data = marketingOrderRepo.findMoAllTypeByMonth(dateMoMonth0, dateMoMonth1, dateMoMonth2);
    	return data;
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
					saveMo.setRevisionPpc(saveMo.getRevisionPpc());
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

    //Show Detail Mo Add PPC di Awal
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
	    
	    //------------------------------------------------------------------------------------------------------------------
	    
	    public List<MonthlyPlan> getAllMp() {
	    	Iterable<MonthlyPlan> mp = monthlyPlanningRepo.findAll();
	        List<MonthlyPlan> mpList = new ArrayList<>();
	        for (MonthlyPlan item : mp) {
	        	MonthlyPlan mpTemp = new MonthlyPlan(item);
	        	mpList.add(mpTemp);
	        }
	        return mpList;
	    }
	    

	    public List<ViewMachineCuring> getMachinesByItemCuring(String itemCuring) {
		    List<Map<String, Object>> result = ctCuringRepo.getMachineByItemCuring(itemCuring);
		    List<ViewMachineCuring> curingList = new ArrayList<>();

		    for (Map<String, Object> map : result) {
		    	ViewMachineCuring curing = new ViewMachineCuring();
		        curing.setOperationShortText((String) map.get("OPERATION_SHORT_TEXT"));
		        
		        curingList.add(curing);
		    }

		    return curingList;
		}
	    
	    public ViewMonthlyPlanning getDetailMonthlyPlan() {
	    	List<Map<String, Object>> result = monthlyPlanningRepo.getDetailMP();
		    List<DetailMonthlyPlanCuring> detailMpl = new ArrayList<>();
		    List<Map<String, Object>> sizePa = new ArrayList<>();
		    List<Map<String, Object>> description = new ArrayList<>();
		    List<DetailDailyMonthlyPlanCuring> detailDailyMpl = new ArrayList<>();

		    for(Map<String, Object> data: result) {
		    	DetailMonthlyPlanCuring detailMp = new DetailMonthlyPlanCuring();
		    	detailMp.setDetailIdCuring((BigDecimal) data.get("DETAIL_ID_CURING"));
		    	detailMp.setDocNumber((String) data.get("DOC_NUM_CURING"));

		    	detailMp.setPartNumber((BigDecimal) data.get("PART_NUMBER"));
		    	detailMp.setTotal((BigDecimal) data.get("TOTAL"));
		    	detailMp.setNetFulfilment((BigDecimal) data.get("NET_FULFILMENT"));
		    	detailMp.setGrossReq((BigDecimal) data.get("GROSS_REQ"));
		       	detailMp.setNetReq((BigDecimal) data.get("NET_REQ"));
		    	detailMp.setReqAhmOem((BigDecimal) data.get("REQ_AHM_OEM"));
		    	detailMp.setReqAhmRem((BigDecimal) data.get("REQ_AHM_REM"));
		    	detailMp.setReqFdr((BigDecimal) data.get("REQ_FDR"));
		    	detailMp.setDifferenceOfs((BigDecimal) data.get("DIFFERENCE_OFS"));
		    	
		    	Map<String, Object> sizee = new HashMap<>();
		    	sizee.put("partNumber", data.get("PART_NUMBER"));
		        sizee.put("description", data.get("DESCRIPTION"));
		        sizee.put("patternName", data.get("PATTERN_NAME"));

		        sizePa.add(sizee);
		    	detailMpl.add(detailMp);
		    	
		    	List<Map<String, Object>> daily2 = monthlyPlanningRepo.getDetailDailyMP((BigDecimal) data.get("DETAIL_ID_CURING"));
			    for(Map<String, Object> data2: daily2) {
			    	DetailDailyMonthlyPlanCuring detailDailyMp = new DetailDailyMonthlyPlanCuring();
			    	detailDailyMp.setDetailDailyIdCuring((BigDecimal) data2.get("DETAIL_DAILY_ID_CURING"));
			    	detailDailyMp.setDetailIdCuring((BigDecimal) data2.get("DETAIL_ID_CURING"));
			    	detailDailyMp.setDateDailyMp((Date) data2.get("DATE_DAILY_MP"));
			    	detailDailyMp.setWorkDay((BigDecimal) data2.get("WORK_DAY"));
			    	detailDailyMp.setTotalPlan((BigDecimal) data2.get("TOTAL_PLAN"));
			    	
			    	Map<String, Object> desc = new HashMap<>();
			    	desc.put("dateDailyMp", data2.get("DATE_DAILY_MP"));
			    	desc.put("description", data2.get("DESCRIPTION"));

			    	description.add(desc);
			    	detailDailyMpl.add(detailDailyMp);
			    	
			    }
		    	
		    }
		    
		    ViewMonthlyPlanning view = new ViewMonthlyPlanning();
		    view.setDetailMonthlyPlanCuring(detailMpl);
		    view.setDetailDailyMonthlyPlanCuring(detailDailyMpl);
		    view.setSizePa(sizePa);
		    view.setDescription(description);

		    return view;
		}
	    
	    public ViewMonthlyPlanning getDetailMonthlyPlanById(String docNum) {
	        MonthlyPlanningCuring monthlyPlanningCuring = monthlyPlanningRepo.findMpById(docNum);
	        
	        List<Map<String, Object>> result = monthlyPlanningRepo.getDetailMPById(docNum);
	        
	        List<DetailMonthlyPlanCuring> detailMpl = new ArrayList<>();
	        List<Map<String, Object>> sizePa = new ArrayList<>();
	        List<Map<String, Object>> description = new ArrayList<>();
	        List<DetailDailyMonthlyPlanCuring> detailDailyMpl = new ArrayList<>();

	        for(Map<String, Object> data: result) {
	            DetailMonthlyPlanCuring detailMp = new DetailMonthlyPlanCuring();
	            detailMp.setDetailIdCuring((BigDecimal) data.get("DETAIL_ID_CURING"));
	            detailMp.setDocNumber((String) data.get("DOC_NUM_CURING"));
	            detailMp.setPartNumber((BigDecimal) data.get("PART_NUMBER"));
	            detailMp.setTotal((BigDecimal) data.get("TOTAL"));
	            detailMp.setNetFulfilment((BigDecimal) data.get("NET_FULFILMENT"));
	            detailMp.setGrossReq((BigDecimal) data.get("GROSS_REQ"));
	            detailMp.setNetReq((BigDecimal) data.get("NET_REQ"));
	            detailMp.setReqAhmOem((BigDecimal) data.get("REQ_AHM_OEM"));
	            detailMp.setReqAhmRem((BigDecimal) data.get("REQ_AHM_REM"));
	            detailMp.setReqFdr((BigDecimal) data.get("REQ_FDR"));
	            detailMp.setDifferenceOfs((BigDecimal) data.get("DIFFERENCE_OFS"));
	            
	            Map<String, Object> sizee = new HashMap<>();
	            sizee.put("partNumber", data.get("PART_NUMBER"));
	            sizee.put("description", data.get("DESCRIPTION"));
	            sizee.put("patternName", data.get("PATTERN_NAME"));
	            sizePa.add(sizee);
	            detailMpl.add(detailMp);
	            
	            List<Map<String, Object>> daily2 = monthlyPlanningRepo.getDetailDailyMP((BigDecimal) data.get("DETAIL_ID_CURING"));
	            for(Map<String, Object> data2: daily2) {
	                DetailDailyMonthlyPlanCuring detailDailyMp = new DetailDailyMonthlyPlanCuring();
	                detailDailyMp.setDetailDailyIdCuring((BigDecimal) data2.get("DETAIL_DAILY_ID_CURING"));
	                detailDailyMp.setDetailIdCuring((BigDecimal) data2.get("DETAIL_ID_CURING"));
	                detailDailyMp.setDateDailyMp((Date) data2.get("DATE_DAILY_MP"));
	                detailDailyMp.setWorkDay((BigDecimal) data2.get("WORK_DAY"));
	                detailDailyMp.setTotalPlan((BigDecimal) data2.get("TOTAL_PLAN"));
	                
	                Map<String, Object> desc = new HashMap<>();
	                desc.put("dateDailyMp", data2.get("DATE_DAILY_MP"));
	                desc.put("description", data2.get("DESCRIPTION"));
	                description.add(desc);
	                detailDailyMpl.add(detailDailyMp);
	            }
	        }
	        
	        ViewMonthlyPlanning view = new ViewMonthlyPlanning();
	        view.setMonthlyPlanningCuring(monthlyPlanningCuring);
	        view.setDetailMonthlyPlanCuring(detailMpl);
	        view.setDetailDailyMonthlyPlanCuring(detailDailyMpl);
	        view.setSizePa(sizePa);
	        view.setDescription(description);
	        
	        return view;
	    }
	    
	    
	    public List<ViewDetailShiftMonthlyPlan> getShiftMonthlyPlan(Date actualDate, BigDecimal detailDailyId ) {
	        List<Map<String, Object>> result = monthlyPlanningRepo.getDetailMonthlyPlan(detailDailyId);
	        if (result == null || result.isEmpty()) {
	            throw new RuntimeException("Detail Monthly Plan tidak ditemukan: " + detailDailyId);
	        }

	        List<ViewDetailShiftMonthlyPlan> detailShiftMpl = new ArrayList<>();

	        for (Map<String, Object> data : result) {
	            List<Map<String, Object>> shift = monthlyPlanningRepo.getDetailShiftMP(actualDate, detailDailyId);
	            for (Map<String, Object> data3 : shift) {
	                ViewDetailShiftMonthlyPlan detailShiftMp = new ViewDetailShiftMonthlyPlan();
	                detailShiftMp.setDetailShiftIdCuring((BigDecimal) data3.get("DETAIL_SHIFT_ID_CURING"));
	                detailShiftMp.setDetailDailyIdCuring((BigDecimal) data3.get("DETAIL_DAILY_ID_CURING"));
	                detailShiftMp.setMoId((String) data3.get("MO_ID"));
	                detailShiftMp.setMpCuringId((String) data3.get("MP_CURING_ID"));
	                detailShiftMp.setActualDate((Date) data3.get("ACTUAL_DATE"));
	                detailShiftMp.setItemCuring((String) data3.get("ITEM_CURING"));
	                detailShiftMp.setPartNumber((BigDecimal) data3.get("PART_NUMBER"));
	                detailShiftMp.setCavity((BigDecimal) data3.get("CAVITY"));
	                detailShiftMp.setCavityUsage((BigDecimal) data3.get("CAVITY_USAGE"));
	                detailShiftMp.setCavityExist((BigDecimal) data3.get("CAVITY_EXIST"));
	                detailShiftMp.setWorkCenterText((String) data3.get("WORK_CENTER_TEXT"));
	                detailShiftMp.setKapaShift1((BigDecimal) data3.get("KAPA_SHIFT_1"));
	                detailShiftMp.setKapaShift2((BigDecimal) data3.get("KAPA_SHIFT_2"));
	                detailShiftMp.setKapaShift3((BigDecimal) data3.get("KAPA_SHIFT_3"));
	                detailShiftMp.setTotalKapa((BigDecimal) data3.get("TOTAL_KAPA"));
	                detailShiftMp.setWct((String) data3.get("WCT"));
	                detailShiftMp.setChangeDate((Date) data3.get("CHANGE_DATE"));
	                detailShiftMp.setShift((BigDecimal) data3.get("SHIFT"));

	                detailShiftMpl.add(detailShiftMp);
	            }
	        }

	        return detailShiftMpl;
	    }

	}

	
