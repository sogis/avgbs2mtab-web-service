package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionDPR {

    public List<String> getOrderedListOfParcelsAffectedByDPRs();

    public List<String> getOrderedListOfNewDPRs();

    public Integer getAddedAreaDPR(String parcelNumberAffectedByDPR, String dpr);

    public Integer getNewAreaDPR(String dpr);

    public Integer getRoundingDifferenceDPR(String dpr);

}
