/*
 * Excel json importer
 */
package exceljsonimport;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import org.json.JSONArray;
import org.json.JSONObject;
import se.toel.app.onyaec.impl.ExcelWriter;
import se.toel.util.Dev;
import se.toel.util.FileUtils;
import se.toel.util.IniFile;

/**
 *
 * @author toel
 */
public class Main {

     private static final String ver="0.1.0";
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        if (args.length<2) showSyntax();
        
        try {
        
            String jsonFileName = args[0];
            String excelFileName = args[1];

            IniFile ini = new IniFile("ExcelJsonImport.ini");
            Map<String, String> conf = new HashMap<>();

            ExcelWriter excel = new ExcelWriter(ini, conf);
            excel.open(excelFileName);
        
            String json = FileUtils.getFileContent(jsonFileName);
            JSONArray rows = new JSONArray(json);
            
            String sheetName = FileUtils.getFileNameWithoutExtention(jsonFileName);
            excel.getWorkbook().createSheet(sheetName);
            excel.selectSheet(sheetName);
            
            Set<String> keys = null;
            String[] data = null;
            for (int i=0; i<rows.length(); i++) {
                JSONObject row = rows.getJSONObject(i);
                
                if (i==0) {
                    keys = row.keySet();
                    excel.setLineValues(keys.toArray(new String[keys.size()]));
                    data = new String[keys.size()];
                }
    
                int n=0;
                for (String key : keys) {
                    data[n++] = row.optString(key);
                }
                excel.setLineValues(data);
                
            }
            
            excel.close();
        
            
        } catch(Exception e) {
            Dev.error("Oops!", e);
        }
        
        
    }
    
    
    private static void showSyntax() {
        
        System.out.println("ExcelJsonImport ver "+ver);
        System.out.println("  Toël Hartmann 2022");
        System.out.println("  Syntax:");
        System.out.println("    java -jar ExcelJsonImport [params] [jsonfile] [excelfile]");
        System.out.println("  where:");
        System.out.println("    [jsonfile] the json file to import");
        System.out.println("    [excelfile] the excel file to update");
        System.exit(1);
        
    }
    
}
