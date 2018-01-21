package ch.so.agi.avgbs2mtab.webservice.services;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Service;

import ch.so.agi.avgbs2mtab.main.Avgbs2mtabMain;

/**
* Spring service class for converting avgbs transfer file to xlsx.
*
* @author  Stefan Ziegler
* @since   2017-08-23
*/
@Service
public class Avgbs2mtabService {
	private final Logger log = LoggerFactory.getLogger(this.getClass());
	
	@Autowired
	private ResourceLoader resourceLoader;
	
	/**
	 * This method converts an AVGBS-INTERLIS transfer file into a xlss file with 
	 * <a href="https://github.com/sogis/avgbs2mtab">avgbs2mtab library</a>.
	 * @param inputFileName Name of AVGBS-INTERLIS transfer file to convert.
	 * @param outputFileName Name of converted xlsx file.
	 * @throws IOException  
	 */	
	public synchronized void convert(String inputFileName, String outFileName) throws IOException {	
		Avgbs2mtabMain.runConversion(inputFileName, outFileName);
	}
}
