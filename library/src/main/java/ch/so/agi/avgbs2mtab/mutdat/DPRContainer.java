package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */


public class DPRContainer implements SetDPR, MetadataOfDPRMutation, DataExtractionDPR {
    Map<String,Map> mainDprMap = new Hashtable<>(); //Main Map. Contains the DPR-Number as key and the laysOnRefAndAreaMap.
    HashMap<String, Integer> laysOnRefAndAreaMap = new HashMap<>(); //Map, containing the Ref of the Parcel and the Area concerning the DPR.
    HashMap<String, String> numberAndRefMap = new HashMap<>(); //Map contains the Ref-String and the Parcel- or DPR-Number.
    HashMap<String, String> affectedParcelsMap = new HashMap<>(); //A small mainParcelMap, containing the numbers of affected Parcels.
    HashMap<String, Integer> newAreaMap = new HashMap<>(); //Map, contains the area of the DPRs.

    ////////////////////////////////////////////////////////
    // SET- Methoden ///////////////////////////////////////
    ///////////////////////////////////////////////////////

    public void setDPRWithAdditions(String dprnumber, String laysonref, Integer area) {
        if(mainDprMap.get(dprnumber) != null) {
            Map laysonrefandarea = mainDprMap.get(dprnumber);
            laysonrefandarea.put(laysonref,area);
        } else {
            laysOnRefAndAreaMap.put(laysonref,area);
        }
        mainDprMap.put(dprnumber, laysOnRefAndAreaMap);
        affectedParcelsMap.put(laysonref,laysonref);
    }

    @Override
    public void setDPRNumberAndRef(String ref, String parcelnumber) {
        numberAndRefMap.put(ref,parcelnumber);
    }

    @Override
    public void setDPRNewArea(String parcelnumber, Integer area) {
        newAreaMap.put(parcelnumber,area);
    }

    ////////////////////////////////////////////////////////
    // GET- Methoden ///////////////////////////////////////
    ///////////////////////////////////////////////////////

    @Override
    public Integer getNumberOfDPRs() {
        int numberofdprs = mainDprMap.size();
        //Addiere noch die gelöschten dazu....
        int i = 0;
        for (String key : newAreaMap.keySet()) {
            if(newAreaMap.get(key).equals(0)) {
                i++;
            }
        }
        numberofdprs += i;
        return numberofdprs;
    }

    @Override
    public Integer getNumberOfParcelsAffectedByDPRs() {
        int numberofparcelsaffected = affectedParcelsMap.size();
        return numberofparcelsaffected;
    }

    @Override
    public List<String> getOrderedListOfParcelsAffectedByDPRs() {
        List<String> parcelsaffectedmydprs = new ArrayList<>();
        List<Integer> ParcelNumbersAsInteger = new ArrayList<>();
        List<String> ParcelNumbersAsString = new ArrayList<>();

        for(String key : affectedParcelsMap.keySet()) {
            String keyparcelnumber = numberAndRefMap.get(key);
            if(!parcelsaffectedmydprs.contains(keyparcelnumber)) {
                parcelsaffectedmydprs.add(keyparcelnumber);
            }
        }

        for (String parcelNumberAsString : parcelsaffectedmydprs)
            ParcelNumbersAsInteger.add(Integer.valueOf(parcelNumberAsString));

        Collections.sort(ParcelNumbersAsInteger);

        for (Integer parcelNumberAsInteger : ParcelNumbersAsInteger)
            ParcelNumbersAsString.add(String.valueOf(parcelNumberAsInteger));

        return ParcelNumbersAsString;
    }

    @Override
    public List<String> getOrderedListOfNewDPRs() {
        List<String> newdprs = new ArrayList<>(mainDprMap.keySet());
        List<Integer> newDPRsAsInteger = new ArrayList<>();
        List<String> newDPRsAsString = new ArrayList<>();

        for (String key : newAreaMap.keySet()) {
            if(newAreaMap.get(key).equals(0)) {
                newdprs.add(key);
            }
        }

        for (String newDPRAsString : newdprs)
            newDPRsAsInteger.add(Integer.valueOf(newDPRAsString));

        Collections.sort(newDPRsAsInteger);

        for (Integer newDPRAsInteger : newDPRsAsInteger)
            newDPRsAsString.add(String.valueOf(newDPRAsInteger));

        return newDPRsAsString;
    }

    @Override
    public Integer getAddedAreaDPR(String parcelNumberAffectedByDPR, String dpr) {
        Map<String,Integer> innerdprmap = mainDprMap.get(dpr);
        String ref = getKeyFromValue(numberAndRefMap,parcelNumberAffectedByDPR);
        Integer addedarea = innerdprmap.get(ref);
        return addedarea;
    }

    @Override
    public Integer getNewAreaDPR(String dpr) {
        int newarea = newAreaMap.get(dpr);
        return newarea;
    }

    @Override
    public Integer getRoundingDifferenceDPR(String dpr) {
        Integer sumaddedareas = 0;
        Integer roundingdifference = null;
        Map<String, Integer> internalmap = mainDprMap.get(dpr);
        if (internalmap != null) {
            for (String key : internalmap.keySet()) {
                Integer area = internalmap.get(key);
                sumaddedareas += area;
            }
            roundingdifference = getNewAreaDPR(dpr) - sumaddedareas;
        }

        return roundingdifference;
    }

    public static String getKeyFromValue(HashMap<String, String> hm, String value) {
        for (String key : hm.keySet()) {
            if (hm.get(key).equals(value)) {
                return key;
            }
        }

        return null;
    }
}
