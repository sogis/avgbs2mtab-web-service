package ch.so.agi.avgbs2mtab.webservice;

import static io.restassured.RestAssured.given;

import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.springframework.boot.context.embedded.LocalServerPort;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.context.SpringBootTest.WebEnvironment;
import org.springframework.test.context.junit4.SpringRunner;

import io.restassured.RestAssured;
import io.restassured.response.Response;

import static org.hamcrest.CoreMatchers.containsString;
import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

import java.io.File;

@RunWith(SpringRunner.class)
@SpringBootTest(webEnvironment = WebEnvironment.RANDOM_PORT)
public class IntegrationTests {
	private final Logger logger = LoggerFactory.getLogger(this.getClass());

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

}
