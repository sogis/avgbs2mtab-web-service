[![Build Status](https://travis-ci.org/sogis/avgbs2mtab-web-service.svg?branch=master)](https://travis-ci.org/sogis/avgbs2mtab-web-service)

# avgbs2mtab-web-service

Spring Boot web service that converts parcel area and ownership mutation data (GB2AV INTERLIS model aka "Kleine Schnittstelle" aka AVGBS) to visual crosstables (Excel file).

## License

avgbs2mtab web service is licensed under the [MIT License](LICENSE).

## Status

avgbs2mtab web service is in production state.

## System Requirements

For the current version of avgbs2mtab web service, you will need a JRE (Java Runtime Environment) installed on your system, version 1.8 or later.

## Developing

avgbs2mtab web service is build as a Spring Boot Application.

`git clone https://github.com/edigonzales/avgbs2mtab-web-service.git` 

Use your favorite IDE (e.g. [Spring Tool Suite](https://spring.io/tools/sts/all)) for coding.

### Testing

`./gradlew clean test` will run all tests: unit tests in the library and functional tests in the web service.

### Building

`./gradlew clean build` will create an executable JAR.

DOCKER...




## TODO

* Test docker image.