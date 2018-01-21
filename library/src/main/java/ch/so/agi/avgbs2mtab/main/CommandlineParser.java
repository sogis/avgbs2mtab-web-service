package ch.so.agi.avgbs2mtab.main;

import org.apache.commons.cli.*;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Defines which arguments can be used as commandline options
 */
public class CommandlineParser {
    private static final Logger log = Logger.getLogger(CommandlineParser.class.getName());
    private String[] args = null;
    private Options options = new Options();

    public CommandlineParser(String[] args) {
        //options = new Options();
        this.args = args;

        options.addOption("h", "help", false, "help");
        options.addOption("o", "output", false, "Outputpath");
        options.addOption("createlog", "createlog", false, "Create Log yes / no");
        options.addOption("log", "logfilepath", false, "Path to the Logfile");
        options.addOption("loglevel", "loglevel", false, "Loglevel Info / Debug / Trace");
        options.addOption("i", "Input", true, "Filepath to XTF Inputfile");
    }

    public String[] parse() {
        CommandLineParser parser = new BasicParser();
        String[] returnArguments = new String[2];
        CommandLine cmd = null;
        try {
            cmd = parser.parse(options, args);

            if (cmd.hasOption("h"))
                help();

            if (cmd.hasOption("o")) {
                log.log(Level.INFO, "The Output file will be " + cmd.getOptionValue("o"));
                returnArguments[0] = cmd.getOptionValue("o");
            }

            if (cmd.hasOption("createlog")) {
                log.log(Level.INFO, "Will create Logfile!");
                returnArguments[1] = cmd.getOptionValue("createlog");
            }

            if ((cmd.hasOption("log")) && (cmd.hasOption("createlog"))){
                log.log(Level.INFO, "Will create Logfile here: "+ cmd.getOptionValue("log"));
                returnArguments[2] = cmd.getOptionValue("log");
            }

            if (cmd.hasOption("loglevel")) {
                log.log(Level.INFO, "Loglevel: "+ cmd.getOptionValue("loglevel"));
                returnArguments[3] = cmd.getOptionValue("loglevel");
            }

            if (cmd.hasOption("i")) {
                log.log(Level.INFO, "Input-File: "+ cmd.getOptionValue("i"));
                returnArguments[4] = cmd.getOptionValue("i");
            }

        } catch (ParseException e) {
            log.log(Level.SEVERE, "Failed to parse command line properties", e);
            help();
        }
        return returnArguments;
    }

    private void help() {
        HelpFormatter formatter = new HelpFormatter();
        formatter.printHelp("Avgbs2mtabMain", options);
        //System.exit(0);
    }
}
