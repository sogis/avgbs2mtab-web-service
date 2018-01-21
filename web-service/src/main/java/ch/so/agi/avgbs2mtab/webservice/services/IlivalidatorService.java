package ch.so.agi.avgbs2mtab.webservice.services;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.interlis2.validator.Validator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Service;

import ch.ehi.basics.settings.Settings;

/**
* Spring service class for INTERLIS transfer file validation.
*
* @author  Stefan Ziegler
* @since   2017-08-23
*/
@Service
public class IlivalidatorService {
	private final Logger log = LoggerFactory.getLogger(this.getClass());
	
	@Autowired
	private Environment env;

	@Autowired
	private ResourceLoader resourceLoader;

	/**
	 * This method validates an INTERLIS transfer file with 
	 * <a href="https://github.com/claeis/ilivalidator">ilivalidator library</a>.
	 * @param inputFileName Name of INTERLIS transfer file.
	 * @param logFileName Name of log file.
	 * @throws IOException if ili file cannot be read or copied to file system. 
	 * @return boolean Returns the validation result.
	 */	
	public synchronized boolean validate(String inputFileName, String logFileName) throws IOException {			
		// Copy model file to our working folder.
		try {
			String modelFileName = env.getProperty("ch.so.agi.avgbs2mtab.webservice.modelFileName"); 

			if (modelFileName == null) {
				log.warn("model file not defined in application.properties");
				modelFileName = "KS3-20060703.ili";
			}

			Resource resource = resourceLoader.getResource("classpath:ili/" + modelFileName);
			InputStream is = resource.getInputStream();

			File iliFile = new File(FilenameUtils.getFullPath(inputFileName), modelFileName);
			Files.copy(is, iliFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
			IOUtils.closeQuietly(is);
		} catch (FileNotFoundException e) {
			log.warn(e.getMessage());
			log.warn("Model file not found. Cannot validate transfer file.");
			throw new IOException(e.getMessage());
		}

		// Validate
		Settings settings = new Settings();
		settings.setValue(Validator.SETTING_ILIDIRS, new File(inputFileName).getParent());
		settings.setValue(Validator.SETTING_LOGFILE, logFileName);
		
		boolean valid = Validator.runValidation(inputFileName, settings);
		
		return valid;
	}	

}
