# About

Reads data provided daily to DfE in .xls format and pushes it into the
[manage courses api](https://github.com/DFE-Digital/manage-courses-api).

# Coding

## ApiClient

`ManageCoursesApiClient.cs` is a client for https://github.com/DFE-Digital/manage-courses-api generated
at build-time from API schema description `manage-courses-api-swagger.json`.

The schema can be updated with `update-schema.sh` or by manually downloading a new copy
to `manage-courses-api-swagger.json` and checking it in.
