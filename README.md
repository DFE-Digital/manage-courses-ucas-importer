# About

Reads data provided daily to DfE in .xls format and pushes it into the
[manage courses api](https://github.com/DFE-Digital/manage-courses-api).

# Build and run

0. Ensure that you are using a version of the ApiClient that is compatible with your deployment target

1. Add the following UCAS data files in the `data/` folder 
```
GTTR_CAMPUS.xls
GTTR_CRSE.xls
GTTR_CRSENOTE.xls
GTTR_CRSE_SUBJECT.xls
GTTR_INST.xls
GTTR_INTAKE.xls
GTTR_NOTETEXT.xls
```

2. Run in the repository root
```
dotnet restore
dotnet build
dotnet run --project src/importer/UcasCourseImporter.csproj --folder data --target <ManageCourses API URL> --key <ManageCourses API admin key>
```