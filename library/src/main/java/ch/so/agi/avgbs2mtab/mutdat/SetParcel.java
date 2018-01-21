package ch.so.agi.avgbs2mtab.mutdat;


public interface SetParcel {


    public void setParcelAddition(String newparcelnumber, String oldparcelnumber, int area);

    public void setParcelNewArea(String newparcelnumber, int newarea);

    public void setParcelOldArea(String oldparcelnumber, int oldarea);

    void delParcelOldArea(String oldparcelnumber);

    public void setParcelRoundingDifference(String parcel, int roundingdifference);

}
