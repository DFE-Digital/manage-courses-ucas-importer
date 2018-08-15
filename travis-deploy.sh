#!/usr/bin/env bash

deployZip=deploy.zip

dotnet publish src/importer --configuration Release

zip $deployZip -r src/importer/bin/Release/netcoreapp2.0/publish

curl -X PUT -u "govuk-dfe-bat-deployment-user:$AZURE_WA_PASSWORD" --data-binary $deployZip https://bat-dev-manage-courses-api-app.scm.azurewebsites.net/api/triggered/ucas-import
