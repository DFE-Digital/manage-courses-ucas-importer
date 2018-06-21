# About

Reads data provided daily to DfE in .xls format and pushes it into the
[manage courses api](https://github.com/DFE-Digital/manage-courses-api).

# Coding

## ApiClient

`ManageCoursesApiClient.cs` is a client for https://github.com/DFE-Digital/manage-courses-api generated
at build-time from API schema description `manage-courses-api-swagger.json`.

The schema can be updated with `update-schema.sh` or by manually downloading a new copy
to `manage-courses-api-swagger.json` and checking it in.

## Usage

`UcasCourseImporter.exe --folder "the folder for all the csv & xls files"`

Expected files in folder
```
GTTR_CAMPUS.xls
GTTR_CRSE.xls
GTTR_CRSENOTE.xls
GTTR_CRSE_SUBJECT.xls
GTTR_INST.xls
GTTR_INTAKE.xls
GTTR_NOTETEXT.xls
GTTR_OUTCOMES.xls
GTTR_PROGRAMME_TYPE.xls
GTTR_REGION.xls
GTTR_SUBJECT.xls
mc-organisations.csv
mc-organisations_institutions.csv
mc-organisations_users.csv
mc-users.csv
ProviderMapper.xls
```