import java.io.*;
import java.util.*;
import com.opencsv.CSVWriter;

public class AuditFileEmployees {

    private static String fileName;
    private static String fileLocation;
    private static File outputAuditFile;
    private static FileWriter outputWriter;
    private static CSVWriter csvWriter;

    public AuditFileEmployees() {

    }

    public AuditFileEmployees(String fileLocation, String fileName) {
        this.fileName = fileName;
        this.fileLocation = fileLocation;
    }

    public void createAuditFile() throws IOException {
        System.out.println("FileLocation:"+fileLocation);
        System.out.println("FileName:"+fileName);
        System.out.println("FL+FN:"+fileLocation+fileName);

        outputAuditFile = new File(fileLocation+fileName);
        outputWriter = new FileWriter(outputAuditFile);
        csvWriter = new CSVWriter(outputWriter);

        String[] header = { "Action", "Field", "Value" };
        csvWriter.writeNext(header);
    }

    public static void writeRow(String action, String field, String value) throws IOException {
        String[] rowDetails = {action,field,value};
        csvWriter.writeNext(rowDetails);
    }

    public static void saveAuditFile() throws IOException {
        csvWriter.close();
    }


}
