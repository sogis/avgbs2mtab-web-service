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

### Release management

It uses the [https://plugins.gradle.org/plugin/pl.allegro.tech.build.axion-release](https://plugins.gradle.org/plugin/pl.allegro.tech.build.axion-release) plugin:

**Condition:** Releases (= Tag on Github, = non-SNAPSHOT version) are made locally.

1. Develop and test and build on your local machine: `./gradlew clean build` 
2. Commit your changes locally: `git commit -a -m 'some fix'`. You cannot make a release without `git push`. Before a release `./gradlew currentVersion` shows `x.y.z-SNAPSHOT`.
3. `./gradlew clean build pushDockerImages` is run on Travis and will push a SNAPSHOT and a latest image on hub.docker.com.
4. If you want a final release (non-SNAPSHOT version), this has to be done locally (commit and push first): `./gradlew release -Prelease.customUsername=foobar -Prelease.customPassword=$*fubarXX! clean build pushDockerImages` (TODO: `release` as last task?). Be carefull: if you push the changes to Github, Travis will be slower with testing and pushing the image than your pushes from your local machine to docker. In this case the `latest` docker image will be overwritten. By default version patch (least significant) number is incremented. This can be changed in `build.gradle` or as command line argument [https://axion-release-plugin.readthedocs.io/en/latest/configuration/version/#incrementing](https://axion-release-plugin.readthedocs.io/en/latest/configuration/version/#incrementing)

## Running as Docker Image (SO!GIS)
* To be done... 

## TODO

* Test docker image.