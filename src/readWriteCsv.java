import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class readWriteCsv {

    public static void csvToXLS(String filename, String sheetno) {
        try {
            String csvFile = filename; //csv file address
            String excelFile = "C:/Users/smitha/Downloads/PJM_V4.xlsx"; //xlsx file address
            HSSFWorkbook workBook = new HSSFWorkbook();
            HSSFSheet sheet = workBook.createSheet(sheetno);
            String currentLine = null;
            int RowNum = 0;
            BufferedReader br = new BufferedReader(new FileReader(csvFile));
            while ((currentLine = br.readLine()) != null) {
                String str[] = currentLine.split(",");
                HSSFRow currentRow = sheet.createRow(RowNum);
                RowNum++;
                for (int i = 0; i < str.length; i++) {
                    currentRow.createCell(i).setCellValue(str[i]);
                }
            }
            FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Done");
        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in try");
        }
    }

    public static List listOfFiles() {
        List<String> results = new ArrayList<String>();
        File[] files = new File("f:\\csv").listFiles();

        for (File file : files) {
            if (file.isFile()) {
                results.add(file.getAbsolutePath());
            }
        }
        return results;
    }

    public static void main(String[] args) {
        List list = listOfFiles();
        int size = list.size();
        for (int i = 0; i < size; i++) {
            csvToXLS(list.get(i).toString(), "sheet" + i + 1);
        }
    }
}