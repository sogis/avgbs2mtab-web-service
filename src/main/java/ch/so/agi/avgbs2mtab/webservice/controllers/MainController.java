package ch.so.agi.avgbs2mtab.webservice.controllers;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.commons.io.FilenameUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import ch.so.agi.avgbs2mtab.webservice.services.Avgbs2mtabService;
import ch.so.agi.avgbs2mtab.webservice.services.IlivalidatorService;


@Controller
public class MainController {
	private final Logger log = LoggerFactory.getLogger(this.getClass());

	// Folder prefix
	private static String FOLDER_PREFIX = "avgbs2mtab_";
    
	@Autowired
	private Environment env;
	
	@Autowired
	private IlivalidatorService ilivalidator;

	@Autowired
	private Avgbs2mtabService avgbs2mtab;
		
	@RequestMapping(value="/", method=RequestMethod.GET)
	public String index() {
		return "index";
	}
	
	@RequestMapping(value = "/", method = RequestMethod.POST)
	@ResponseBody
	public ResponseEntity<?> uploadFile(
			@RequestParam(name="file", required=true) MultipartFile uploadFile
			) 
	{				
		try {
			// Get the filename and build the local file path.
			String filename = uploadFile.getOriginalFilename();
			String directory = env.getProperty("ch.so.agi.avgbs2mtab.webservice.uploadedFiles"); 

			if (directory == null) {
				directory = System.getProperty("java.io.tmpdir");
			}
			
			Path tmpDirectory = Files.createTempDirectory(Paths.get(directory), FOLDER_PREFIX);
			Path uploadFilePath = Paths.get(tmpDirectory.toString(), filename);

			// Save the file locally.			
			byte[] bytes = uploadFile.getBytes();
			Files.write(uploadFilePath, bytes);
			
			log.debug("uploadFilePath: " + uploadFilePath);
			
			// Validate avgbs transfer file.
			String inputFileName = uploadFilePath.toString();
			String baseFileName = FilenameUtils.getFullPath(inputFileName) 
					+ FilenameUtils.getBaseName(inputFileName);
			String logFileName = baseFileName + "_validator.log";

			boolean valid = ilivalidator.validate(inputFileName, logFileName);
			
			log.debug("Is transfer file valid: " + valid);
			
			if (!valid) {
				File logFile = new File(logFileName);
				InputStream is = new FileInputStream(logFile);

				return ResponseEntity
				.ok()
				.contentLength(logFile.length())
				.contentType(MediaType.parseMediaType("text/plain"))
				.body(new InputStreamResource(is));	    
			}
			
			// Convert avgbs transfer file into a xlsx file.
			String outputFileName = baseFileName + ".xlsx";

			avgbs2mtab.convert(inputFileName, outputFileName);
			
		
			// Send log file back to client.
//			File logFile = new File(logFileName);
//			InputStream is = new FileInputStream(logFile);
//
//			return ResponseEntity
//					.ok()
//					.contentLength(logFile.length())
//					.contentType(MediaType.parseMediaType("text/plain"))
//					.body(new InputStreamResource(is));	    
			
			return ResponseEntity
					.ok().body("adf");
		}
		catch (Exception e) {
			e.printStackTrace();
			log.error(e.getMessage());
			return ResponseEntity
					.badRequest()
					.contentType(MediaType.parseMediaType("text/plain"))
					.body(e.getMessage());
		}

	} 

}
