FROM bellsoft/liberica-openjdk-alpine-musl:17.0.1-12

EXPOSE 8878

WORKDIR /home/app

COPY build/libs/web-service*.jar /home/app/application.jar

RUN chown -R 1001:0 /home/app && \
    chmod -R g=u /home/app

USER 1001

ENV LOG4J_FORMAT_MSG_NO_LOOKUPS=true

ENTRYPOINT ["java","-Djava.security.egd=file:/dev/./urandom","-jar","application.jar"]