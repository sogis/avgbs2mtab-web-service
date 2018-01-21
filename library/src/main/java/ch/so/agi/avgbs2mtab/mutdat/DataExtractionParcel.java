package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionParcel {

    public List<String> getOrderedListOfOldParcelNumbers();

    public List<String> getOrderedListOfNewParcelNumbers();

    public Integer getAddedArea(String oldParcelNumber, String newParcelNumber);

    public Integer getNewArea(String newParcelNumber);

    public Integer getRoundingDifference(String oldParcelNumber);

    public Integer getRestAreaOfParcel(String oldParcelNumber);

}
