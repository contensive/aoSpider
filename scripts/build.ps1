#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\contensive-build.psm1') -Force

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

Invoke-ContensiveBuild `
    -CollectionName    'Spider' `
    -CollectionPath    "$projectRoot\collections\aoSpider" `
    -SolutionPath      "$projectRoot\server\Spider.sln" `
    -BinPath           "$projectRoot\server\Spider\bin\Release" `
    -DeploymentRoot    'C:\Deployments\aoSpider' `
    -CleanFolders      @(
                           "$projectRoot\server\Spider\bin"
                           "$projectRoot\server\Spider\obj"
                       )
