package com.belwar.aiseexcelprocessor;

import java.util.ArrayList;

public class PatologyBlock {

    public int total = 0;
    private Integer[][] patologies = new Integer[][] { new Integer[] { new Integer (0) },
                                                       new Integer[] { new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0) },                                      
                                                       new Integer[] { new Integer (0), new Integer (0), new Integer (0), 
                                                                       new Integer (0), new Integer (0) },
                                                       new Integer[] { new Integer (0) } };
    
    public void setPatology (String idStr, int ammount) {

        String [] auxString = idStr.split("\\.");
        
        try {
            int id, subId;   
                 
            id = Integer.parseInt(auxString[0]) - 1;
            
            if (auxString.length < 2 || auxString[1] == null)
                subId = 0;
            else 
                subId = Integer.parseInt(auxString[1]);
            patologies[id][subId] = new Integer(ammount);            
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(0);
        }
    }
    
    public void computeTotal() {
        int newTotal = 0;
        for (Integer[] ia : patologies) {
            for (Integer i : ia) {
                newTotal += i.intValue();
            }
        }
        total = newTotal;
    }
    
    public ArrayList<Integer> retrievePatologies(){
    
        ArrayList<Integer> patologies = new ArrayList<Integer>();
        
        for (Integer[] ia : this.patologies) {
            for (Integer i : ia) {
                patologies.add(i);
            }
        }
        
        return patologies;
    }
}
