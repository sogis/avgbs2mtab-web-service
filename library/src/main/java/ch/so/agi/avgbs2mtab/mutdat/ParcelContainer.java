package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */
public class ParcelContainer implements SetParcel, MetadataOfParcelMutation, DataExtractionParcel {

    Map<String,Map> mainParcelMap = new Hashtable<>(); //The main parcel-map.
    Map<String,Integer> parcelNewAreaMap = new Hashtable<>(); //Map contains the new area of a parcel.
    Map<String,Integer> parcelRemainingAreaMap = new Hashtable<>(); //Map contains the remaining-area of a Parcel (Diagonale).
    Map<String,Integer> parcelRoundingDifferenceMap = new Hashtable<>(); //Map contains the roundingdifference.

    ///////////////////////////////////////////////
    // SET-Methoden //////////////////////////////
    //////////////////////////////////////////////

    @Override
    public void setParcelAddition(String newparcelnumber, String oldparcelnumber, int area) {

        Map<String,Integer> parcelmap = mainParcelMap.get(newparcelnumber);
        if(parcelmap == null){
            parcelmap = new Hashtable<String, Integer>();
            mainParcelMap.put(newparcelnumber, parcelmap);
        }
        parcelmap.put(oldparcelnumber, area);
    }

    @Override
    public void setParcelNewArea(String newparcelnumber, int newarea) {
            parcelNewAreaMap.put(newparcelnumber,newarea);
    }

    @Override
    public void setParcelOldArea(String oldparcelnumber, int oldarea) {
            parcelRemainingAreaMap.put(oldparcelnumber,oldarea);
    }

    @Override
    public void delParcelOldArea(String oldparcelnumber) {
        parcelRemainingAreaMap.remove(oldparcelnumber);
    }

    @Override
    public void setParcelRoundingDifference(String parcel, int roundingdifference) {
            parcelRoundingDifferenceMap.put(parcel,roundingdifference);
    }

    ////////////////////////////////////
    //GET-Methoden  ///////////////////
    ///////////////////////////////////

    @Override
    public List<String> getOrderedListOfOldParcelNumbers() {
        List<String> oldparcelnumbers = new ArrayList<>();
        //Add all parcelnumbers in the inner-mainDprMap from the main mainDprMap to the oldparcelmap
        for(String key : mainParcelMap.keySet()) {
            Map internalmap = mainParcelMap.get(key);
            for(Object keyoldparcels : internalmap.keySet()) {
                if(!oldparcelnumbers.contains(keyoldparcels)) {
                    oldparcelnumbers.add((String) keyoldparcels);
                }
            }
        }
        //Add also all parcelnumbers from parcelRemainingAreaMap to the oldparcelmap
        List<Integer> oldParcelNumbersAsInteger = new ArrayList<>();
        List<String> oldParcelNumbersAsString = new ArrayList<>();

        for(String key : parcelRemainingAreaMap.keySet()) {
            if(!oldparcelnumbers.contains(key)) {
                oldparcelnumbers.add(key);
            }
        }

        for (String parcelNumberAsString : oldparcelnumbers)
            oldParcelNumbersAsInteger.add(Integer.valueOf(parcelNumberAsString));

        Collections.sort(oldParcelNumbersAsInteger);

        for (Integer parcelNumberAsInteger : oldParcelNumbersAsInteger)
            oldParcelNumbersAsString.add(Integer.toString(parcelNumberAsInteger));

        return oldParcelNumbersAsString;
    }

    @Override
    public List<String> getOrderedListOfNewParcelNumbers() {
        List<String> newparcelnumbers = new ArrayList<>(parcelNewAreaMap.keySet());
        List<Integer> newParcelNumbersAsInteger = new ArrayList<>();
        List<String> newParcelNumbersAsString = new ArrayList<>();

        for (String parcelNumberAsString : newparcelnumbers)
            newParcelNumbersAsInteger.add(Integer.valueOf(parcelNumberAsString));

        Collections.sort(newParcelNumbersAsInteger);

        for (Integer parcelNumberAsInteger : newParcelNumbersAsInteger)
            newParcelNumbersAsString.add(String.valueOf(parcelNumberAsInteger));

        return newParcelNumbersAsString;
    }
    @Override
    public Integer getAddedArea(String newparcel, String oldparcel) {
        Map addmap = mainParcelMap.get(newparcel);
        Integer areaadded = null; //Etwas unglücklich, aber bisher unvermeidlich!
        if (addmap!=null) {
            areaadded = (Integer) addmap.get(oldparcel);
        }
        return areaadded;
    }

    @Override
    public Integer getNewArea(String newParcelNumber) {
        Integer newarea = parcelNewAreaMap.get(newParcelNumber);
        return newarea;
    }

    @Override
    public Integer getRoundingDifference(String oldParcelNumber) {
        Integer roundingdifference = parcelRoundingDifferenceMap.get(oldParcelNumber);
        return roundingdifference;
    }

    @Override
    public Integer getNumberOfOldParcels() {
        Integer numberofoldparcels = getOrderedListOfOldParcelNumbers().size();
        return numberofoldparcels;
    }

    @Override
    public Integer getNumberOfNewParcels() {
        Integer numberofnewparcels = getOrderedListOfNewParcelNumbers().size();
        return numberofnewparcels;
    }

    @Override
    public Integer getRestAreaOfParcel(String oldParcelNumber) {
        Integer restarea = parcelRemainingAreaMap.get(oldParcelNumber);
        return restarea;
    }


}
