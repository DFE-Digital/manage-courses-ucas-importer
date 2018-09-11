

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

2. Set config options

```bash
# .\src\importer>

# refer to https://github.com/DFE-Digital/manage-courses-api
dotnet user-secrets set manage_api_url the-manage-api-url (ie. http://localhost:6001)
dotnet user-secrets set manage_api_key the-manage-api-key (ie. the same value set for https://github.com/DFE-Digital/manage-courses-api "refer to api:key" ) 

# values available from portal.azure.com 
dotnet user-secrets set azure_url the-azure-url
dotnet user-secrets set azure_signature the-azure-signature
```


3. Run in the repository root
```bash
# .\src\importer>
dotnet restore
dotnet build
dotnet run
```

## Logging

Logging is configured in `appsettings.json`, and values in there can be overridden with environment variables.

Powershell:

    $env:Serilog:MinimumLevel="Debug"
    dotnet run

Command prompt

    set Serilog:MinimumLevel=Debug
    dotnet run

For more information see:

* https://github.com/serilog/serilog-settings-configuration
* https://nblumhardt.com/2016/07/serilog-2-minimumlevel-override/

Serilog has been configured to spit logs out to both the console
(for `dotnet run` testing & development locally) and Application Insights.

Set the `APPINSIGHTS_INSTRUMENTATIONKEY` environment variable to tell Serilog the application insights key.
