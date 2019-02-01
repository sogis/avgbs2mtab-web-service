# avgbs2mtab-web-service

## First install in an Openshift Environment

All necessary components of the application are configured in the template avgbs2mtab-web-service.yaml
Set environment with parameter env and docker image version with parameter version
```
oc process -p env=test -p version=latest -f avgbs2mtab-web-service.yaml  | oc apply -f-
```

## Update of app configuration in Openshift Environment

Make changes to the configuration in the template avgbs2mtab-web-service.yaml and run
Set environment with parameter env and docker image version with parameter version
```
oc process -p env=test -p version=latest -f avgbs2mtab-web-service.yaml  | oc apply -f-
```
