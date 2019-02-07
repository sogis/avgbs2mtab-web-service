package ch.so.agi.avgbs2mtab.webservice;

import static io.restassured.RestAssured.given;

import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.context.SpringBootTest.WebEnvironment;
import org.springframework.boot.web.server.LocalServerPort;
import org.springframework.test.context.junit4.SpringRunner;

import io.restassured.RestAssured;

import static org.hamcrest.CoreMatchers.containsString;
import static org.hamcrest.CoreMatchers.equalTo;

import java.io.File;
import java.io.IOException;

@RunWith(SpringRunner.class)
@SpringBootTest(webEnvironment = WebEnvironment.RANDOM_PORT)
public class IntegrationTests {
	private final Logger logger = LoggerFactory.getLogger(this.getClass());

	@Rule
	public TemporaryFolder tempFolder = new TemporaryFolder();

    @LocalServerPort
    int randomServerPort;
    
    @Before
    public void setPort() {
    	RestAssured.port = randomServerPort;
    }
	
	@Test
	public void indexPageTest() {				
		given().
		when().
        	get("/avgbs2mtab/").
        then().
        	statusCode(200).
        	body("html.head.title", equalTo("avgbs2mtab web service"));
	}
	
    /*
     * Test if version.txt is available.
     */
    @Test
    public void versionPageTest() {               
        given().
        when().
            get("/avgbs2mtab/version.txt").
        then().
            statusCode(200).
            body(containsString("Revision")).
            body(containsString("Application-name"));
    }

	@Test
	/*
	 * Sending an empty file is like pushing the upload
	 * button without selecting a file first.
	 */
	public void emptyFileUploadTest() throws IOException {
		File file = tempFolder.newFile("tempFile.txt");

		given().
			multiPart("file", file).
		when().
			post("/avgbs2mtab/").
		then().
	    	statusCode(302);
	}
	
	@Test
	/*
	 * Upload a non-AVGBS-INTERLIS file.
	 */
	public void wrongInterlisFileUploadTest() {
		File file = new File("src/test/data/agglo_20170529.xtf");

		given().
			multiPart("file", file).
		when().
			post("/avgbs2mtab/").
		then().
	    	statusCode(200).
	    	body(containsString("model not found"));		
	}	
	
	@Test
	/*
	 * Upload a text file with nonsense content.
	 */		
	public void nonsenseFileUploadTest() {
		File file = new File("src/test/data/nonsense.txt");

		given().
			multiPart("file", file).
		when().
			post("/avgbs2mtab/").
		then().
	    	statusCode(200).
	    	body(containsString("model not found"));
	}

	@Test
	/*
	 * Upload an AVGBS file with errors.
	 * INTERLIS validation must fail.
	 */
	public void avgbsFileWithErrorsTest() {
		File file = new File("src/test/data/SO0200002406_999_20161116_errors.xml");

		given().
			multiPart("file", file).
		when().
			post("/avgbs2mtab/").
		then().
	    	statusCode(200).
	    	body(containsString("tid z48f364f300002077: Attribute Nummer requires a value")).
	    	body(containsString("Info: ...validation failed"));
	}
	
	@Test
	/*
	 * Upload an AVGBS file and a XSLX will
	 * be returned (= successful conversion).
	 * 
	 * There seems to be a bug with ContentType.BINARY: 
	 * https://github.com/rest-assured/rest-assured/issues/861
	 * 
	 * Use headers instead. Be careful with content length.
	 * This can (?) change when using different xlsx 
	 * library/version.
	 */
	public void successfulConversionTest() {
		File file = new File("src/test/data/SO0200002402_970_20161222.xml");
		
		given().
			multiPart("file", file).
		when().
			post("/avgbs2mtab/").
		then().
	    	statusCode(200).
	    	header("Content-Type", "application/octet-stream").
	    	header("content-disposition", "attachment; filename=SO0200002402_970_20161222.xlsx");
	    	//header("Content-Length","5174"); // failed even with new avgbs2mtab version w/o changing xlsx lib.
	}
}
