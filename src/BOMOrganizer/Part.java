package BOMOrganizer;

import java.util.ArrayList;

/**
 * Created by rahul on 2017-08-15.
 */
public class Part { //used to store info of unique parts
    private ArrayList<String> designators;
    private String partNumber;
    private ArrayList<String> otherInfo;

    Part(ArrayList<String> des, String pn, ArrayList<String> otherIn){
        designators = des;
        partNumber = pn;
        otherInfo = otherIn;
    }

    public String getPartNumber() {
        return partNumber;
    }

    public ArrayList<String> getDesignators() {

        return designators;
    }

    public ArrayList<String> getOtherInfo() {
        return otherInfo;
    }
}
