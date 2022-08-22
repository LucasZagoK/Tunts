package tunts;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

@SuppressWarnings("rawtypes")
public class App {
    private static Sheets sheetsService;    // Access SDK
    private static String APPLICATION_NAME = "Tunts";
    private static String SPREADSHEET_ID = "1FM2xFisoaV_PS0P5LHq2LBJngqnVcODvVlanZbJbNBg";
    private static Logger log = Logger.getLogger("Log");

    // Grant app access to google sheets
    private static Credential authorize() throws IOException, GeneralSecurityException {
        // Load client secrets.
        InputStream in = App.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(GsonFactory.getDefaultInstance(), new InputStreamReader(in));

        // List of scopes
        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        // Give Application access to spreadsheets
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
        GoogleNetHttpTransport.newTrustedTransport(), GsonFactory.getDefaultInstance(), clientSecrets, scopes)
            .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
            .setAccessType("offline")
            .build();
        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");

        return credential;
    }

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        Credential credential = authorize();
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), GsonFactory.getDefaultInstance(), credential)
            .setApplicationName(APPLICATION_NAME)
            .build();
    }
    // Return all content from the sheet
    public static void printSheet(List<List<Object>> values) throws IOException{
        System.out.print("[");
        for (List row : values) {
            System.out.print(" {");
            for (Object item : row) {
                if (row.indexOf(item) == row.size()-1) {
                    System.out.print(item);
                    continue;
                }
                System.out.print(item + " - ");
            }
            System.out.println("}");
        }
        System.out.print("]");
    }

    // Return spreadsheet data in a specific range
    public static List<List<Object>> getValues(String range) throws IOException{
        // Returns a list of lists with the values in the range especified
        ValueRange response = sheetsService.spreadsheets().values()
            .get(SPREADSHEET_ID, range)
            .execute();
        
        List<List<Object>> items = response.getValues();
        
        return items;
    }

    // Returns an int with total number of classes
    public static int getTotalClasses(List<List<Object>> values){
        int totalClasses = 0;
        String item = "";
        for (List row : values) {
            item = (String)row.get(0);
            break;
        }
            char[] chars = item.toCharArray();
            StringBuilder sb = new StringBuilder();
            for (char c : chars) {
                if (Character.isDigit(c)) {
                    sb.append(c);
                }
            }
            item = sb.toString();
            totalClasses = Integer.parseInt(item);
        
        return totalClasses;
    }

    // Update spreadsheet values on G4H27 range  
    public static void updateValues(ValueRange body, String range) throws IOException, GeneralSecurityException{
        sheetsService.spreadsheets().values().update(SPREADSHEET_ID, range, body)
        .setValueInputOption("RAW")
        .execute();
    }

    // Return specific student average grade 
    public static float getAverageGrade(float p1, float p2, float p3) {
        float m = (p1+p2+p3)/3;
        return m;
    }

    // Return "Nota para aprovacao final"
    public static float getNAF(float m) {
        float naf = 100 - m;
        return naf;
    }

    // Initialize and generate a file for logs
    public static void initLogger() {
        FileHandler fh;
        try {
            fh = new FileHandler(System.getProperty("user.dir")+"/ActivityLog.log");

            log.addHandler(fh);
            SimpleFormatter formatter = new SimpleFormatter();
            fh.setFormatter(formatter);

            log.info("Logger initialized");
        } catch (Exception e) {
            log.log(Level.WARNING, "Exeption: ",e);
        }
    }

    public static void main(String[] args) throws IOException, GeneralSecurityException{
        initLogger();   // Initialize logger
        try {
            sheetsService = getSheetsService();
            log.info("Access to sheets granted");
        } catch (Exception e) {
            log.log(Level.WARNING, "Exeption: ",e);
        }
        
        String range = "engenharia_de_software!A2:H27";
        
        // Values that will be on G4 to H27 
        ValueRange body = new ValueRange(); 
        
        // Extract data from spreadsheet
        List<List<Object>> values = getValues(range);
        try {
            printSheet(values);
            log.info("Successfully printed data");            
        } catch (Exception e) {
            log.log(Level.WARNING, "Exeption: ",e);
        }

        int totalClasses = getTotalClasses(values);
        int limit = (totalClasses * 25)/100;

        range = "engenharia_de_software!C4:H27";
        values = getValues(range);
         
        //  Read useful data from spreadsheet to set the data on G4H27 range        
        for (List row : values) {
            int faults = Integer.parseInt(row.get(0).toString());

            //  Check if the student is "Reprovado por falta" or not
            if (faults > limit) { 
                //  Set G column value
                body.setValues(Arrays.asList(Arrays.asList("Reprovado por falta")));
                updateValues(body, "G"+Integer.toString(values.lastIndexOf(row)+4));

                //  Set H column value
                body.setValues(Arrays.asList(Arrays.asList("0")));
                updateValues(body, "H"+Integer.toString(values.lastIndexOf(row)+4));       

                log.info("Successfully update student");         
                continue;
            }

            //  Grades from students
            float p1 = Float.parseFloat(row.get(1).toString());
            float p2 = Float.parseFloat(row.get(2).toString());
            float p3 = Float.parseFloat(row.get(3).toString());
            float m = getAverageGrade(p1, p2, p3);

            //  Check students "Situacao" and "Nota para aprovacao final" and set those values
            if (m >= 70) {
                try {
                    //  Set G column value
                    body.setValues(Arrays.asList(Arrays.asList("Aprovado")));
                    updateValues(body, "G"+Integer.toString(values.lastIndexOf(row)+4));
    
                    //  Set H column value
                    body.setValues(Arrays.asList(Arrays.asList("0")));
                    updateValues(body, "H"+Integer.toString(values.lastIndexOf(row)+4));    
                    
                    log.info("Successfully update student");
                    
                } catch (Exception e) {
                    log.log(Level.WARNING, "Exeption: ",e);
                }
                continue;

            } else if (m >= 50) {
                try {
                    //  Set G column value
                    body.setValues(Arrays.asList(Arrays.asList("Exame Final")));
                    updateValues(body, "G"+Integer.toString(values.lastIndexOf(row)+4));
    
                    //  Set H column value with naf
                    body.setValues(Arrays.asList(Arrays.asList(Float.toString(Math.round(getNAF(m))))));
                    updateValues(body, "H"+Integer.toString(values.lastIndexOf(row)+4));       
                    
                    log.info("Successfully update student");    
                } catch (Exception e) {
                    log.log(Level.WARNING, "Exeption: ",e);
                }
                continue;
            }
            try {
                body.setValues(Arrays.asList(Arrays.asList("Reprovado por Nota")));
                updateValues(body, "G"+Integer.toString(values.lastIndexOf(row)+4));
    
                //  Set H column value with naf
                body.setValues(Arrays.asList(Arrays.asList("0")));
                updateValues(body, "H"+Integer.toString(values.lastIndexOf(row)+4));
                log.info("Successfully update student");
            } catch (Exception e) {
                log.log(Level.WARNING, "Exeption: ",e);
            }
        }

        values = getValues("engenharia_de_software!A2:H27");
        try {
            printSheet(values);
            log.info("Successfully printed data");            
        } catch (Exception e) {
            log.log(Level.WARNING, "Exeption: ",e);
        }

        log.info("End of Program");
    }

}
