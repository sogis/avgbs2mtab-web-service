package ch.so.agi.avgbs2mtab.mutdat;

/**
 *
 */
public interface SetDPR {

    public void setDPRWithAdditions(String dprnumber, String laysonref, Integer area);

    void setDPRNumberAndRef(String ref, String parcelnumber);

    void setDPRNewArea(String parcelnumber, Integer area);

}
