# About

Reads data provided daily to DfE in .xls format and pushes it into the
[manage courses api](https://github.com/DFE-Digital/manage-courses-api).

# Build and run

1. Download the swagger.json of the API you want to deploy to. This is found at http://<host>/swagger/v1/swagger.json, e.g. https://manage-courses-api-bat-development.e4ff.pro-eu-west-1.openshiftapps.com/swagger/v1/swagger.json - save this as as `src/api-client/manage-courses-api-swagger.json`

2. Add the following data files in the `data/` folder 
```
#From UCAS data dump:
GTTR_CAMPUS.xls
GTTR_CRSE.xls
GTTR_CRSENOTE.xls
GTTR_CRSE_SUBJECT.xls
GTTR_INST.xls
GTTR_INTAKE.xls
GTTR_NOTETEXT.xls

#From DTTP import script 
mc-organisations.csv
mc-organisations_institutions.csv
mc-organisations_users.csv
mc-users.csv
```

3. Run in the repository root
```
dotnet restore
dotnet build
dotnet run --project src/importer/UcasCourseImporter.csproj --folder data
```