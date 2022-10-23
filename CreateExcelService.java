import com.sun.media.sound.InvalidFormatException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.apache.commons.io.FileUtils;
import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

public class CreateExcelService {

    @Autowired
    private ServletContext servletContext;

    private static FileInputStream inputStream;
    private static Workbook workbook = new XSSFWorkbook();

    private static HashMap<Integer, String> rowName = new HashMap<Integer, String>(){{
        put(2,"12");
        put(2,"13");
        put(2,"14");
        put(2,"15");
        put(2,"16");
        put(2,"17");
        put(2,"18");
        put(2,"19");
        put(2,"20");
        put(2,"21");
        put(2,"22");
        put(2,"23");
        put(2,"24");
        put(2,"25");
    }};

    public CreateExcelService()throws IOException, InvalidFormatException {

    }

    public void createExcel(HttpServletResponse response, String year, String month){


        String excelFilePath = servletContext.getRealPath("/assets/excelReportFiles/.xlsx");
        String outputExcelPath = servletContext.getRealPath("/assets/excelReportFiles/.xlsx");

        try{
            inputStream = new FileInputStream(new File(excelFilePath));
            workbook = WorkbookFactory.create(inputStream);
            rowName.keySet().forEach(e->{
                updateRow(e,rowName.get(e),year,month);
            });

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            formulaEvaluator.evaluateAll();


            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(outputExcelPath);
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            System.out.println("finish......");
            getFile(response);

        }catch (IOException | EncryptedDocumentException | org.apache.poi.openxml4j.exceptions.InvalidFormatException exception){
            exception.getStackTrace();
        }
    }
    public void getFile(HttpServletResponse response){
        String outputExcelPath = servletContext.getRealPath("/assets/exportReportFiles/xlsx");

        try{
            FileInputStream file   = new FileInputStream( new File(outputExcelPath));
            byte[] bytes = FileUtils.readFileToByteArray(new File(outputExcelPath));
            OutputStream os = response.getOutputStream();
            response.setContentType("appliation/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setContentLength(bytes.length);
            os.write(bytes);
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    public void  updateRow(int rowNum, String statusCode, String year, String month){
        Sheet mainSheet = workbook.getSheet("گزراش");
        org.apache.poi.ss.usermodel.Row row = ((Sheet) mainSheet).getRow(rowNum);
        RetiredGeneralReportExcelDto  r = iPersonnelRepostitory.getRetiredPersonnelReport(statusCode, year, month);
        if (r == null){
            row.getCell(2).setCellValue(0);
            row.getCell(3).setCellValue(0D);
            row.getCell(4).setCellValue(0);
            row.getCell(5).setCellValue(0);
        }else{
            row.getCell(2).setCellValue(r.getTotalCount());
            row.getCell(3).setCellValue(r.getSumMembershipFee());
            row.getCell(4).setCellValue(r.getNotPayCount());
            row.getCell(5).setCellValue(r.getPayCount());
        }
        createRowDetail(rowNum, statusCode, year, month);
    }

    public void linkToSheet(Row row){
        int rowNum = row.getRowNum();
        CreationHelper creationHelper = workbook.getCreationHelper();
        Hyperlink hyperlink  = creationHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
        hyperlink.setAddress("'" + rowName.get(rowNum) + "!A1");
        Cell cell = row.getCell(1);
        cell.setHyperlink(hyperlink);
    }
    public  void createRowDetail(int rowNum, String statusCode, String year, String month){

        String sheetName = rowName.get(rowNum);
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.setRightToLeft(true);
        org.apache.poi.ss.usermodel.Row headerRow =  sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();

        headerRow.createCell(0).setCellValue("ردیف");
        headerRow.createCell(1).setCellValue("شماره شناسایی");
        headerRow.createCell(2).setCellValue("شناسه ملی");
        headerRow.createCell(3).setCellValue("نام و نام خانوادگی");
        headerRow.createCell(4).setCellValue("مبلغ عضویت کانون");
        headerRow.createCell(5).setCellValue("شناسه وضعیت");
        headerRow.createCell(6).setCellValue("شرح وضعیت");


        List<RetiredGeneralReportExcelDto> retiredPersonnel = iPerosnnelService.getRetiredPersonnel(statusCode, year, month);

        AtomicInteger rowIndex = new AtomicInteger(1);

        retiredPersonnel.forEach(e->{
            org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowIndex.get());
            row.createCell(0).setCellValue(e.getPersonCode());
            row.createCell(0).setCellValue(e.getNationalCode());
            row.createCell(0).setCellValue(e.getFullName());
            row.createCell(0).setCellValue(e.getMembershipFee());
            row.createCell(0).setCellValue(e.getStatusCode());
            row.createCell(0).setCellValue(e.getStatusDescription());
        });

        for (int i = 0; i < 7; i++) {
            headerRow.getCell(i).setCellStyle(cellStyle);
        }

        sheet.setAutoFilter(new CellRangeAddress(0,0,0,6));
        sheet.createFreezePane(0,1);
    }

    public void download(HttpServletResponse response){

    }

    public boolean isReportExist(){
        String filePath = servletContext.getRealPath("/assets/excelReportFiles/.xlsx");
        File file = new File(filePath);
        return file.exists();
    }

}
