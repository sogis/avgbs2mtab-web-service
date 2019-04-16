#!/usr/bin/env groovy

pipeline {
    agent any

    stages {
        stage('Prepare') {
            steps {
                echo "Checkout sources"
                git "https://github.com/sogis/avgbs2mtab-web-service.git/"
            }
        }
        
        stage('Build') {
            steps {
                sh './gradlew --no-daemon clean classes'
            }
        }

        stage('Test') {
            steps {
                sh './gradlew --no-daemon test'
                publishHTML target: [
                    reportName : 'Web Service Tests',
                    reportDir:   'web-service/build/reports/tests/test', 
                    reportFiles: 'index.html',
                    keepAll: true,
                    alwaysLinkToLastBuild: true,
                    allowMissing: false
                ]                
            }
        }

        stage('Publish') {
            steps {
                sh './gradlew --no-daemon bootJar'  
                archiveArtifacts artifacts: "web-service/build/libs/web-service-*.jar", onlyIfSuccessful: true, fingerprint: true                              
            }
        }               
    }
    post {
        always {
            deleteDir() 
        }
    }
}