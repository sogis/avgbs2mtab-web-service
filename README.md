[![CI/CD](https://github.com/sogis/avgbs2mtab-web-service/actions/workflows/build.yml/badge.svg)](https://github.com/sogis/avgbs2mtab-web-service/actions/workflows/build.yml)

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

## Building

### JVM

```
./gradlew clean build
```

### Native

```
./gradlew clean build

java -DspringAot=true -agentlib:native-image-agent=config-output-dir=web-service/src/main/resources/META-INF/native-image -jar web-service/build/libs/web-service-1.1.LOCALBUILD-exec.jar

./gradlew clean aotTest nativeCompile -i
```

## Run

```

```

Es ist keine Anwendungs-Konfiguration vorzunehmen, d.h. keine umgebungsspezifische Properties-Dateien oder ENV-Variablen.


## Release management / versioning

It uses a simple release management and versioning mechanism: Local builds are tagged as `1.0.LOCALBUILD`. Builds on Github Action will append the build number, e.g. `1.0.48`. Major version will be increased after "major" changes. After every commit to the repository a docker image will be build and pushed to `hub.docker.com`. It will be tagged as `latest` and with the build number (`1.0.48`).


## TODO

* Test docker image.


