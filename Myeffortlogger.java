package excelWriterExample;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Myeffortlogger {
     public static void main1(String[] args) {
            FileOutputStream csp = null;
            try {
                csp = new FileOutputStream(new File("csp of Effort Logger.xlsx"));
                
                String SAMPLE_XLSX_FILE_PATH = "170040_A_AppDev_W10_EffortLogger.xlsx";
                Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
                workbook.write(csp);
                workbook.close();
                csp.flush();
            }
            catch (Exception d) {
                d.printStackTrace();
            }
     }
}
