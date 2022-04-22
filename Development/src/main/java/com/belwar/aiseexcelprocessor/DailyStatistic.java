package com.belwar.aiseexcelprocessor;

import java.time.LocalDate;
import java.util.ArrayList;

public class DailyStatistic implements Comparable<DailyStatistic> {
    public String groupName;
    public LocalDate date;
    public String town;
    public String place;
    
    public ArrayList<TalkBlock> talks;
    
    public PatologyBlock girlsPatologies;
    public PatologyBlock womenPatologies;
    public PatologyBlock menPatologies;
    public PatologyBlock boysPatologies;
    
    public String comments;
    
    @Override
    public int compareTo(DailyStatistic ds) {
        return date.compareTo(ds.date);
    }
}
