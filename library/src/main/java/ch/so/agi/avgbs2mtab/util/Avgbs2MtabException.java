package ch.so.agi.avgbs2mtab.util;

public class Avgbs2MtabException  extends RuntimeException {

        public static final String TYPE_NO_FILE = "TYPE_NO_FILE";
        public static final String TYPE_WRONG_EXTENSION = "TYPE_WRONG_EXTENSION";
        public static final String TYPE_NO_XML_STYLING = "TYPE_NO_XML_STYLING";
        public static final String TYPE_TRANSFERDATA_NOT_FOR_AVGBS_MODEL = "TYPE_TRANSFERDATA_NOT_FOR_AVGBS_MODEL";
        public static final String TYPE_FILE_NOT_READABLE = "TYPE_FILE_NOT_READABLE";
        public static final String TYPE_FOLDER_NOT_WRITEABLE = "TYPE_FOLDER_NOT_WRITEABLE";
        public static final String TYPE_FILE_EXISTS = "TYPE_FILE_EXISTS";
        public static final String TYPE_VALIDATION_FAILED = "TYPE_VALIDATION_FAILED";
        public static final String TYPE_MISSING_PARCEL_IN_EXCEL = "TYPE_MISSING_PARCEL_IN_EXCEL";
        public static final String TYPE_NUMBERFORMAT = "TYPE NUMBERFORMAT";

        private String type;

        public Avgbs2MtabException(){}

        public Avgbs2MtabException(String message) {
            super(message);
        }

        public Avgbs2MtabException(String message, Throwable cause) {
            super(message, cause);
        }

        public Avgbs2MtabException(String type, String message){
            super(message);
            this.type = type;

        }

        public String getType(){
            return this.type;
        }
}
