import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

import com.opencsv.CSVReader;

public class ImportDropDownLists {

    public ImportDropDownLists() {

    }

    public ArrayList OpenDropDownCSV(String fileLocation){

        CSVReader reader = null;
        ArrayList<String> businessUnits = new ArrayList<String>();
        try {
            //Get the CSVReader instance with specifying the delimiter to be used
            reader = new CSVReader(new FileReader(fileLocation),';');
            String [] nextLine;
            //Read one line at a time
            while ((nextLine = reader.readNext()) != null)
            {
                for(String token : nextLine)
                {
                    //Print all tokens
//                    System.out.println(token);
                    businessUnits.add(token.toString());
                }
            }
        }
        catch (Exception e) {
            e.printStackTrace();
            return null;
        }
        finally {
            try {
                reader.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            return  businessUnits;
        }
    }

}
