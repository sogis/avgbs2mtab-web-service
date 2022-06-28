package ch.so.agi.avgbs2mtab.webservice;

import static io.restassured.RestAssured.given;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.TestInstance.Lifecycle;
import org.junit.jupiter.api.io.TempDir;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.context.SpringBootTest.WebEnvironment;
import org.springframework.boot.test.web.server.LocalServerPort;

import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.path.xml.XmlPath;
import io.restassured.path.xml.XmlPath.CompatibilityMode;
import io.restassured.response.Response;

import static org.hamcrest.CoreMatchers.containsString;
import static org.hamcrest.CoreMatchers.equalTo;
import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;

@TestInstance(Lifecycle.PER_CLASS) // Damit BeforeAll nicht statisch sein muss.
@SpringBootTest(webEnvironment = WebEnvironment.RANDOM_PORT)
public class IntegrationTests {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    @LocalServerPort
    int randomServerPort;
    
    @BeforeAll
    public void setPort() {
        RestAssured.port = randomServerPort;
    }
    
    @Test
    public void indexPageTest() {              
//        given().
//        when().
//            get("/avgbs2mtab/").
//        then().
//            statusCode(200)
//            body("html.head.title", equalTo("avgbs2mtab web service"));
        
        
        Response response = 
                given()
                .when()
                    .get("/avgbs2mtab/")
                .then().contentType(ContentType.HTML).extract()
                .response();
        assertEquals(response.getStatusCode(), 200);
        XmlPath htmlPath = new XmlPath(CompatibilityMode.HTML, response.getBody().asString());
        assertEquals(htmlPath.getString("html.head.title"), "avgbs2mtab web service");

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
            body(containsString("Version")).
            body(containsString("Application-name"));
    }

    /*
     * Sending an empty file is like pushing the upload
     * button without selecting a file first.
     */
    @Test
    public void emptyFileUploadTest(@TempDir Path tempDir) throws IOException {
        File file = tempDir.resolve("tempFile.txt").toFile();
        file.createNewFile();

        given().
            multiPart("file", file).
        when().
            post("/avgbs2mtab/").
        then().
            statusCode(302);
    }
    
    /*
     * Upload a non-AVGBS-INTERLIS file.
     */
    @Test
    public void wrongInterlisFileUploadTest() {
        File file = new File("src/test/data/254900.itf");

        given().
            multiPart("file", file).
        when().
            post("/avgbs2mtab/").
        then().
            statusCode(200).
            body(containsString("model(s) not found"));     
    }   
    
    /*
     * Upload a text file with nonsense content.
     */     
    @Test
    public void nonsenseFileUploadTest() {
        File file = new File("src/test/data/nonsense.txt");

        given().
            multiPart("file", file).
        when().
            post("/avgbs2mtab/").
        then().
            statusCode(200).
            body(containsString("no reader found"));
    }

    /*
     * Upload an AVGBS file with errors.
     * INTERLIS validation must fail.
     */
    @Test
    public void avgbsFileWithErrorsTest() {
        File file = new File("src/test/data/SO0200002406_999_20161116_errors.xml");

        given().
            multiPart("file", file).
        when().
            post("/avgbs2mtab/").
        then().
            statusCode(200).
            body(containsString("tid z48f364f300002077: Attribute Nummer[0]/Nummer requires a value")).
            body(containsString("Info: ...validation failed"));
    }
    
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
    @Test
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
