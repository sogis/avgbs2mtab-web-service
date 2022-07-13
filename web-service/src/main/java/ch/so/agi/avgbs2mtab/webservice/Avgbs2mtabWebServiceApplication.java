package ch.so.agi.avgbs2mtab.webservice;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
//import org.springframework.nativex.hint.TypeHint;
//import org.springframework.nativex.hint.TypeAccess;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.filter.ForwardedHeaderFilter;

//@TypeHint(
//        types = {org.apache.xmlbeans.impl.store.Locale.class},
//        access= {TypeAccess.DECLARED_METHODS, 
//              TypeAccess.DECLARED_FIELDS, 
//              TypeAccess.DECLARED_CONSTRUCTORS, 
//              TypeAccess.PUBLIC_METHODS,
//              TypeAccess.PUBLIC_FIELDS,
//              TypeAccess.PUBLIC_CONSTRUCTORS}               
//)
@Configuration
@SpringBootApplication
public class Avgbs2mtabWebServiceApplication {

	public static void main(String[] args) {
		SpringApplication.run(Avgbs2mtabWebServiceApplication.class, args);
	}
	
    @Bean
    ForwardedHeaderFilter forwardedHeaderFilter() {
        return new ForwardedHeaderFilter();
    }  
}
