package ch.so.agi.avgbs2mtab.webservice.controllers;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import javax.servlet.ServletContext;

import org.apache.commons.io.FilenameUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
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
    private ServletContext servletContext;

	@Autowired
	private IlivalidatorService ilivalidator;

	@Autowired
	private Avgbs2mtabService avgbs2mtab;
		
	@RequestMapping(value="/", method=RequestMethod.GET)
	public String index() {
		return "index";
	}
    
	@RequestMapping(value = "/version.txt", method = RequestMethod.GET)
    public String version() {
        return "version.txt";
    }
    
	@RequestMapping(value = "/", method = RequestMethod.POST)
	@ResponseBody
	public ResponseEntity<?> uploadFile(
			@RequestParam(name="file", required=true) MultipartFile uploadFile
			) {		
		
		try {
			// Get the filename.
			// Need to use FilenameUtils.getName() since getOriginalFilename() returns absolute path for files sent with IE (on macOS only?)
			String filename = FilenameUtils.getName(uploadFile.getOriginalFilename());
			File file = new File(filename);
			
			// If the upload button was pushed w/o choosing a file,
			// we just redirect to the starting page.
			if (uploadFile.getSize() == 0 
					|| filename.trim().equalsIgnoreCase("")
					|| filename == null) {
				log.warn("No file was uploaded. Redirecting to starting page.");
				
				HttpHeaders headers = new HttpHeaders();
				headers.add("Location", servletContext.getContextPath());    
				return new ResponseEntity<String>(headers, HttpStatus.FOUND);			
			}			
			
			// Build the local file path.
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
			
			// Send xlsx file back to client.
			File xlsxFile = new File(outputFileName);
			InputStream is = new FileInputStream(xlsxFile);

			return ResponseEntity
					.ok().header("content-disposition", "attachment; filename=" + xlsxFile.getName())
					.contentLength(xlsxFile.length())
					.contentType(MediaType.parseMediaType("application/octet-stream"))
					.body(new InputStreamResource(is));	    
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
