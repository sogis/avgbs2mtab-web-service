package ch.so.agi.avgbs2mtab.readxtf;

import ch.interlis.iom.IomObject;
import ch.interlis.iom_j.xtf.XtfReader;
import ch.interlis.iox.*;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.function.Consumer;
import java.util.logging.Level;
import java.util.logging.Logger;

import ch.so.agi.avgbs2mtab.mutdat.DataExtractionParcel;
import ch.so.agi.avgbs2mtab.mutdat.SetDPR;
import ch.so.agi.avgbs2mtab.mutdat.SetParcel;
import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;


/**
 * This Class contains methods to read xtf-files and write specific content to a hashtable
 */
public class ReadXtf {

    private static final Logger LOGGER = Logger.getLogger(ReadXtf.class.getName());
    private static final String ILI_MODELNAME ="GB2AV";
    private final String ILI_MUT= ILI_MODELNAME +".Mutationstabelle";
    private IoxReader ioxReader=null;
    private SetParcel parceldump;
    private SetDPR drpdump;

    private HashSet<String> parcelMetadataSet = null; //contains all the Refs that are affected in this Mutation.
    private HashMap<String,HashMap> dprMetadataMap = null; //contains the metadata of the drps.


    public ReadXtf(SetParcel parceldump, SetDPR drpdump) {

        this.parceldump = parceldump;
        this.drpdump = drpdump;
    }

    public void readFile(String xtfFilePathString) throws IOException {
        LOGGER.log(Level.CONFIG,"Start reading the transferfile with multiple passes");

        Consumer<StartBasketEvent> getParcelMetadataFunc = this::getParcelMetadata;
        Consumer<StartBasketEvent> getDRPMetadataFunc  = this::getDRPMetadata;
        Consumer<StartBasketEvent> transferParcelFunc = this::transferParcel;
        Consumer<StartBasketEvent> transferDPRFunc = this::transferDPR;

        File xtfFilePath = new File(xtfFilePathString);

        parseTransferFile(xtfFilePath, getParcelMetadataFunc, true);
        parseTransferFile(xtfFilePath, getDRPMetadataFunc, false);
        parseTransferFile(xtfFilePath, transferParcelFunc, false);
        parseTransferFile(xtfFilePath, transferDPRFunc, false);
    }

    private void parseTransferFile(File xtfFilePath, Consumer innerFunction, boolean checkModelMatch){
        try{
            // open xml file
            ioxReader= new XtfReader(xtfFilePath);
            // loop threw baskets
            IoxEvent event;
            while(true){
                event=ioxReader.read();;
                if(event instanceof ObjectEvent){
                }else if(event instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event;

                    if(checkModelMatch){
                        assertModelIsAvGbs(se);
                    }

                    innerFunction.accept(se);

                }else if(event instanceof EndBasketEvent){
                }else if(event instanceof StartTransferEvent){
                    StartTransferEvent se=(StartTransferEvent)event;
                    String sender=se.getSender();
                }else if(event instanceof EndTransferEvent){
                    System.out.flush();
                    ioxReader.close();
                    ioxReader=null;
                    break;
                }
            }
        }catch(Exception e){
            if(e instanceof Avgbs2MtabException) {
                throw (Avgbs2MtabException) e;
            }
            else {
                throw new Avgbs2MtabException("Error in reading xtf file for innerFunction " + innerFunction, e);
            }
        }
        finally{
            if(ioxReader!=null){
                try{
                    ioxReader.close();
                }catch(IoxException ex){
                    throw new Avgbs2MtabException("Error closing IoxReader", ex);
                }
                ioxReader=null;
            }
        }
    }


    private void getParcelMetadata(StartBasketEvent basket) {
        // loop through basket and find all "betroffeneGrundstuecke"
        HashSet<String> oidBetroffeneGrundstuecke = new HashSet<>();
        try {
            IoxEvent event;
            while (true) {
                event = ioxReader.read();
                if (event instanceof ObjectEvent) {
                    IomObject iomObj = ((ObjectEvent) event).getIomObject();
                    String aclass = iomObj.getobjecttag();

                    if (aclass.equals(ILI_MUT + ".AVMutationBetroffeneGrundstuecke")) {
                        String ref = iomObj.getattrobj("betroffeneGrundstuecke", 0).getobjectrefoid();

                        oidBetroffeneGrundstuecke.add(ref);
                    }
                } else if (event instanceof EndBasketEvent) {
                    break;
                } else {
                    throw new IllegalStateException("unexpected event " + event.getClass().getName());
                }
            }
        }
        catch (IoxException ix){
            throw new Avgbs2MtabException("IoxException in getParcelMetadata", ix);
        }
        this.parcelMetadataSet = oidBetroffeneGrundstuecke;
    }


    private void getDRPMetadata(StartBasketEvent basket) {
        // loop threw basket and find all "betroffeneGrundstuecke"

        HashMap<String,HashMap> dprAnteilAnLiegenschaft = new HashMap<String, HashMap>();

        try {
            HashMap<String, Integer> liegtaufmap = new HashMap<String, Integer>();
            IoxEvent event;
            while (true) {
                event = ioxReader.read();
                if (event instanceof ObjectEvent) {
                    IomObject iomObj = ((ObjectEvent) event).getIomObject();
                    String aclass = iomObj.getobjecttag();

                    if (aclass.equals(ILI_MODELNAME + ".Grundstuecksbeschrieb.Anteil")) {
                        String drpnumber = iomObj.getattrobj("flaeche", 0).getobjectrefoid();
                        String liegt_auf = iomObj.getattrobj("liegt_auf", 0).getobjectrefoid();
                        int area = (int) (Double.parseDouble(iomObj.getattrvalue("Flaechenmass"))*10);

                        if (dprAnteilAnLiegenschaft.get(drpnumber) != null) {
                            liegtaufmap = dprAnteilAnLiegenschaft.get(drpnumber);
                            liegtaufmap.put(liegt_auf, area);
                        } else {
                            liegtaufmap.put(liegt_auf, area);
                        }
                        dprAnteilAnLiegenschaft.put(drpnumber, liegtaufmap);

                    }
                } else if (event instanceof EndBasketEvent) {
                    break;
                } else {
                    throw new IllegalStateException("unexpected event " + event.getClass().getName());
                }
            }
        }
        catch (IoxException ix){
            throw new Avgbs2MtabException("IoxException in getParcelMetadata", ix);
        }
        this.dprMetadataMap = dprAnteilAnLiegenschaft;
    }

    private void transferParcel(StartBasketEvent basket) {
        try {
            //loop threw basket, find things and write them to the Container
            IoxEvent event2;
            while(true){
                event2=ioxReader.read();
                if(event2 instanceof ObjectEvent){
                    IomObject iomObj=((ObjectEvent)event2).getIomObject();
                    String aclass=iomObj.getobjecttag();
                    if(aclass.equals(ILI_MUT+".Liegenschaft")){
                        getParcelLiegenschaft(iomObj);
                    }
                }else if(event2 instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event2.getClass().getName());
                }
            }
        }
        catch (IoxException ix){
            throw new Avgbs2MtabException("IoxException thrown", ix);
        }
    }

    private void getParcelLiegenschaft(IomObject iomObj) {
        if(iomObj.getattrvalue("GrundstueckArt").equals("Liegenschaft")) {
            if (parcelMetadataSet.contains(iomObj.getobjectoid())) {
                String parcelnumber = iomObj.getattrobj("Nummer", 0).getattrvalue("Nummer");
                int area = (int) (Double.parseDouble(iomObj.getattrvalue("Flaechenmass"))*10);
                parceldump.setParcelNewArea(parcelnumber, area);
                parceldump.setParcelOldArea(parcelnumber,area); //vorläufig wird die alte Fläche = neue Fläche gesetzt.

                String roundingDiffString = iomObj.getattrvalue("Korrektur");
                if(roundingDiffString != null && roundingDiffString.length() > 0){
                    int roundingDifference = Integer.parseInt(roundingDiffString);
                    parceldump.setParcelRoundingDifference(parcelnumber, roundingDifference);
                }

                getZugaenge(iomObj, parcelnumber, area);
            }
        }
    }

    private void getZugaenge(IomObject iomObj, String parcelnumber, int area) {
        if (iomObj.getattrvaluecount("Zugang") > 0) {
            int additionsum = 0;
            for (int i = 0; i < iomObj.getattrvaluecount("Zugang"); i++) {
                String oldparcelnumber = iomObj.getattrobj("Zugang", i).getattrobj("von", 0).getattrvalue("Nummer");
                int additionarea = (int) (Double.parseDouble(iomObj.getattrobj("Zugang", i).getattrvalue("Flaechenmass"))*10);
                parceldump.setParcelAddition(parcelnumber, oldparcelnumber, additionarea);
                additionsum += additionarea;
            }
            if(area != additionsum){
                parceldump.setParcelOldArea(parcelnumber,area-additionsum);
            }
            else{
                parceldump.delParcelOldArea(parcelnumber);
            }

        }
    }

    /////////////////////////////////////
    //Get DPRData    ////////////////////
    /////////////////////////////////////

    private void transferDPR(StartBasketEvent se) {

        try {
            //loop threw basket, find things and write them to the Container
            IoxEvent event;
            while (true) {
                event = ioxReader.read();
                if (event instanceof ObjectEvent) {
                    IomObject iomObj = ((ObjectEvent) event).getIomObject();
                    String aclass = iomObj.getobjecttag();
                    getDPR(dprMetadataMap, iomObj, aclass);
                    getDPRLiegenschaft(iomObj, aclass);
                    getDeletedDPR(iomObj, aclass);
                }else if(event instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event.getClass().getName());
                }
            }
        }catch(IoxException ex){
            LOGGER.log(Level.WARNING,"Got an Error in transferDPR: "+ex);
            throw new Avgbs2MtabException("Error in transferDPR, see inner Exception", ex);
        }
    }

    private void getDeletedDPR(IomObject iomObj, String aclass) {
        if (aclass.equals(ILI_MUT + ".AVMutation")) {
            Integer numberofdeletedparcels = iomObj.getattrvaluecount("geloeschteGrundstuecke");
            for(Integer i=0;i<numberofdeletedparcels;++i) {
                String nummer = iomObj.getattrobj("geloeschteGrundstuecke", i).getattrvalue("Nummer");
                DataExtractionParcel parceldatagetter = (DataExtractionParcel)parceldump;
                if(!parceldatagetter.getOrderedListOfOldParcelNumbers().contains(nummer)&&!parceldatagetter.getOrderedListOfNewParcelNumbers().contains(nummer)) {
                    drpdump.setDPRNewArea(nummer, 0);
                }
            }
        }
    }

    private void getDPRLiegenschaft(IomObject iomObj, String aclass) {
        if (aclass.equals(ILI_MUT + ".Liegenschaft")) {
            Integer numbercount = iomObj.getattrvaluecount("Nummer");
            for (int i = 0;i<numbercount;++i) {
                String parcelnumber = iomObj.getattrobj("Nummer", i).getattrvalue("Nummer");
                String parcelref = iomObj.getobjectoid();
                drpdump.setDPRNumberAndRef(parcelref, parcelnumber);
            }
        }
    }

    private void getDPR(HashMap<String, HashMap> drpmetadatamap, IomObject iomObj, String aclass) {
        if (aclass.equals(ILI_MUT + ".Flaeche")) {
            if (iomObj.getattrvalue("GrundstueckArt").startsWith("SelbstRecht")) { //Selbstrecht.* ?
                if (drpmetadatamap.containsKey(iomObj.getobjectoid())) {
                    String parcelnumber = iomObj.getattrobj("Nummer", 0).getattrvalue("Nummer");
                    int newarea = (int) (Double.parseDouble(iomObj.getattrvalue("Flaechenmass"))*10);
                    String parcelref = iomObj.getobjectoid();
                    Map internalmap = drpmetadatamap.get(iomObj.getobjectoid());
                    for (Object key : internalmap.keySet()) {
                        String fromparcelref = key.toString();
                        Integer area = Integer.parseInt(internalmap.get(key).toString());
                        drpdump.setDPRWithAdditions(parcelnumber,fromparcelref,area);
                    }
                    drpdump.setDPRNumberAndRef(parcelref,parcelnumber);
                    drpdump.setDPRNewArea(parcelnumber,newarea);
                }
            }
        }
    }

    /////////////////////////////////////
    //Utility        ////////////////////
    /////////////////////////////////////

    private static void assertModelIsAvGbs(StartBasketEvent se){

        String namev[] = se.getType().split("\\.");
        String modelName = namev[0];

        if(!modelName.equals(ILI_MODELNAME)) {
            throw new Avgbs2MtabException(
                    Avgbs2MtabException.TYPE_TRANSFERDATA_NOT_FOR_AVGBS_MODEL,
                    String.format("Given transferfile references wrong model %s (should reference %s)", modelName, ILI_MODELNAME)
            );
        }
    }
}
