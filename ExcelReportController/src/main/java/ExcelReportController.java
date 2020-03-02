import java.io.IOException;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


@Path("/ajaxReport")
public class ExcelReportController {

    public static Logger logger = LoggerFactory.getLogger(ExcelReportController.class);

    private static final String password = "foobar1";

    @GET
    @Path("{attribute1}/{attribute2}")
    @Produces("application/vnd.ms-excel")
    public Response getExcelDownload(@PathParam("attribute1") String attribute1,
            @PathParam("attribute2") String attribute2)
                    throws IOException, InvalidFormatException, GeneralSecurityException {

        HashMap<String, Object> map = new HashMap<>();
        List<String> headers = new ArrayList<>();
        headers.add("Header 1");
        headers.add("Header 2");
        headers.add("Header 3");
        List<Map<String, String>> valueList = new ArrayList<>();
        valueList.add(map1);
        map.put("header", headers);
        map.put("values", valueList);

        XSSFWorkbook workbook = new XSSFWorkbook();
        OPCPackage opc = workbook.getPackage();

        exportExcel(workbook, map);

        StreamingOutput fileStream = new StreamingOutput() {
            @Override
            public void write(java.io.OutputStream output) throws IOException, WebApplicationException {
                try {
                    // For enc
                    createAndWriteEncryptedWorkbook(workbook, opc, output);
                    workbook.write(output);

                } catch (Exception e) {
                }
            }
        };

        ResponseBuilder response = Response.ok(fileStream,
                "application/vnd.ms-excelformats-officedocument.spreadsheetml.sheet");
        response.header("content-disposition", "attachment; filename=" + System.currentTimeMillis() + ".xlsx");
        return response.build();
    }

    public static void exportExcel(XSSFWorkbook workbook, HashMap<String, Object> map) throws IOException {

        XSSFSheet sheet = workbook.createSheet("Report Sheet");
        List<String> headers = (List<String>) map.get("header");
        List<Map<String, String>> valueList = (List<Map<String, String>>) map.get("values");
        // Create a blank sheet
        XSSFRow rowHead = sheet.createRow((short) 0);
        for (int i = 0; i < headers.size(); i++) {
            rowHead.createCell((short) i).setCellValue(headers.get(i));
        }
        int rownum = 0;
        for (int j = 0; j < valueList.size(); j++) {
            rownum = j + 1;
            XSSFRow row = sheet.createRow(rownum);
            Map<String, String> valueMap = valueList.get(j);
            Set<Map.Entry<String, String>> entrySet = valueMap.entrySet();
            int colNum = 0;
            for (Entry<String, String> entry : entrySet) {
                Cell cellFName = row.createCell(colNum, Cell.CELL_TYPE_STRING);
                cellFName.setCellValue(entry.getValue());
                colNum++;
            }
        }
        sheet.protectSheet(password);

        return workbook;
    }

    private void createAndWriteEncryptedWorkbook(XSSFWorkbook workbook, OPCPackage opc,
            OutputStream requestOutputStream) throws IOException {
        // populateWorkbook(workbook);

        try {
            // Add password protection and encrypt the file
            POIFSFileSystem fileSystem = new POIFSFileSystem();
            opc.save(getEncryptingOutputStream(fileSystem, password));
            fileSystem.writeFilesystem(requestOutputStream);
        } finally {
            // workbook.close();
        }
    }

    private OutputStream getEncryptingOutputStream(POIFSFileSystem fileSystem, String password) throws IOException {
        EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
        Encryptor encryptor = encryptionInfo.getEncryptor();
        encryptor.confirmPassword(password);

        try {
            return encryptor.getDataStream(fileSystem);
        } catch (GeneralSecurityException e) {
            // TODO handle this better
            throw new RuntimeException(e);
        }
    } 
}