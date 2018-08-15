#!/usr/bin/env bash
dotnet publish src/importer --configuration Release
zip deploy.zip -r src/importer/bin/Release/netcoreapp2.0/publish
curl -X PUT --data deploy.zip --header "Content-Disposition: attachment; attachment; filename=Copy.zip" --header "Authorization: " https://b.scm.azurewebsites.net/api/triggered/ucas-import