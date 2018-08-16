#!/usr/bin/env bash

deployZip=deploy.zip

dotnet publish src/importer --configuration Release

zip $deployZip -r src/importer/bin/Release/netcoreapp2.0/publish

curl -X DELETE -u "$1" https://bat-dev-manage-courses-api-app.scm.azurewebsites.net/api/triggeredwebjobs/ucas-import
curl -X PUT -u "$1" --data-binary $deployZip --header "Content-Type: application/zip" --header "Content-Disposition: attachment; filename=deploy.zip" https://bat-dev-manage-courses-api-app.scm.azurewebsites.net/api/triggeredwebjobs/ucas-import/
