package sri.sysint.sri_starter_back.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import sri.sysint.sri_starter_back.model.Plant;
import sri.sysint.sri_starter_back.repository.PlantRepo;


@Service
@Transactional
public class PlantServiceImpl {
	@Autowired
    private PlantRepo plantRepo;
	
    public PlantServiceImpl(PlantRepo plantRepo){
        this.plantRepo = plantRepo;
    }
    
    public BigDecimal getNewId() {
    	return plantRepo.getNewId().add(BigDecimal.valueOf(1));
    }

    public List<Plant> getAllPlant() {
    	Iterable<Plant> plants = plantRepo.getDataOrderId();
        List<Plant> plantList = new ArrayList<>();
        for (Plant item : plants) {
            Plant plantTemp = new Plant(item);
            plantList.add(plantTemp);
        }
        
        return plantList;
    }
    
    public Optional<Plant> getPlantById(BigDecimal id) {
    	Optional<Plant> plant = plantRepo.findById(id);
    	return plant;
    }
    
    public Plant savePlant(Plant plant) {
        try {
        	plant.setPLANT_ID(getNewId());
            plant.setSTATUS(BigDecimal.valueOf(1));
            plant.setCREATION_DATE(new Date());
            plant.setLAST_UPDATE_DATE(new Date());
            return plantRepo.save(plant);
        } catch (Exception e) {
            System.err.println("Error saving plant: " + e.getMessage());
            throw e;
        }
    }
    
    public Plant updatePlant(Plant plant) {
        try {
            Optional<Plant> currentPlantOpt = plantRepo.findById(plant.getPLANT_ID());
            
            if (currentPlantOpt.isPresent()) {
                Plant currentPlant = currentPlantOpt.get();
                
                currentPlant.setPLANT_NAME(plant.getPLANT_NAME());
                currentPlant.setLAST_UPDATE_DATE(new Date());
                currentPlant.setLAST_UPDATED_BY(plant.getLAST_UPDATED_BY());
                
                return plantRepo.save(currentPlant);
            } else {
                throw new RuntimeException("Plant with ID " + plant.getPLANT_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating plant: " + e.getMessage());
            throw e;
        }
    }
    
    public Plant deletePlant(Plant plant) {
        try {
            Optional<Plant> currentPlantOpt = plantRepo.findById(plant.getPLANT_ID());
            
            if (currentPlantOpt.isPresent()) {
                Plant currentPlant = currentPlantOpt.get();
                
                currentPlant.setSTATUS(BigDecimal.valueOf(0));
                currentPlant.setLAST_UPDATE_DATE(new Date());
                currentPlant.setLAST_UPDATED_BY(plant.getLAST_UPDATED_BY());
                
                return plantRepo.save(currentPlant);
            } else {
                throw new RuntimeException("Plant with ID " + plant.getPLANT_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error updating plant: " + e.getMessage());
            throw e;
        }
    }
    
    public Plant restorePlant(Plant plant) {
        try {
            Optional<Plant> currentPlantOpt = plantRepo.findById(plant.getPLANT_ID());
            
            if (currentPlantOpt.isPresent()) {
                Plant currentPlant = currentPlantOpt.get();
                
                currentPlant.setSTATUS(BigDecimal.valueOf(1)); 
                currentPlant.setLAST_UPDATE_DATE(new Date());
                currentPlant.setLAST_UPDATED_BY(plant.getLAST_UPDATED_BY());
                
                return plantRepo.save(currentPlant);
            } else {
                throw new RuntimeException("Plant with ID " + plant.getPLANT_ID() + " not found.");
            }
        } catch (Exception e) {
            System.err.println("Error restoring plant: " + e.getMessage());
            throw e;
        }
    }

    
    public void deleteAllPlant() {
    	plantRepo.deleteAll();
    }
    
    
    public ByteArrayInputStream exportPlantsExcel () throws IOException {
    	List<Plant> plants = plantRepo.getDataOrderId();
    	ByteArrayInputStream byteArrayInputStream = dataToExcel(plants);
    	return byteArrayInputStream;
    }
    
    public ByteArrayInputStream dataToExcel(List<Plant> plants) throws IOException {
        String[] header = {
            "NOMOR",
            "PLANT_ID",
            "PLANT_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PLANT DATA");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setAlignment(HorizontalAlignment.CENTER); 

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER); 

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 1;
            int nomor = 1;
            for (Plant p : plants) {
                Row dataRow = sheet.createRow(rowIndex++);

                Cell nomorCell = dataRow.createCell(0);
                nomorCell.setCellValue(nomor++);
                nomorCell.setCellStyle(borderStyle);

                Cell idCell = dataRow.createCell(1);
                idCell.setCellValue(p.getPLANT_ID().doubleValue());
                idCell.setCellStyle(borderStyle);

                Cell nameCell = dataRow.createCell(2);
                nameCell.setCellValue(p.getPLANT_NAME());
                nameCell.setCellStyle(borderStyle);
            }

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to export data");
            return null;
        } finally {
            workbook.close();
            out.close();
        }
    }

    public ByteArrayInputStream layoutPlantsExcel () throws IOException {
    	ByteArrayInputStream byteArrayInputStream = layoutPlantExcel();
    	return byteArrayInputStream;
    }
    
    public ByteArrayInputStream layoutPlantExcel() throws IOException {
        String[] header = {
            "NOMOR",
            "PLANT_ID",
            "PLANT_NAME"
        };

        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            Sheet sheet = workbook.createSheet("PLANT DATA");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            borderStyle.setAlignment(HorizontalAlignment.CENTER);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderStyle);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                cell.setCellStyle(headerStyle);
            }

            for (int rowIndex = 1; rowIndex <= 3; rowIndex++) {
                Row dataRow = sheet.createRow(rowIndex);
                for (int colIndex = 0; colIndex < header.length; colIndex++) {
                    Cell cell = dataRow.createCell(colIndex);
                    cell.setCellStyle(borderStyle);
                }
            }

            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to create layout");
            throw e;
        } finally {
            workbook.close();
            out.close();
        }
    }
}