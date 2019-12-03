<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.
Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$SiteUrl,
    [Parameter(Mandatory=$true)]
    [String]$UserName,
    [switch]$DiagOnly
)

# Load the required assemblies. 
# Note that SharePoint Online CSOM 16.1.8361.1200 is the required version of this sample.
$PathOfSPClientDll = "C:\csom\lib\net45\Microsoft.SharePoint.Client.dll";
$PathOfSPClientRuntimeDll = "C:\csom\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"

Add-Type -Path $PathOfSPClientDll;
Add-Type -Path $PathOfSPClientRuntimeDll;
$assemblies = ([System.Reflection.AssemblyName]::GetAssemblyName($PathOfSPClientDll).FullName, [System.Reflection.AssemblyName]::GetAssemblyName($PathOfSPClientRuntimeDll).FullName);

function ExecuteQueryWithIncrementalRetry
{
    param (
        [int]$retryCount,
        [int]$delay = 120
    );

    $RetryAfterHeaderName = "Retry-After";
    $retryAttempts = 0;
    $backoffInterval = $delay
    $retryAfterInterval = 0;
    $retry = $false;

    if ($retryCount -le 0) {
        throw "Provide a retry count greater than zero.";
    }
    if ($delay -le 0) {
        throw "Provide a delay greater than zero.";
    }

    while ($retryAttempts -lt $retryCount) {
        try {
            if (!$retry)
            {
                $script:context.ExecuteQuery();
                return;
            }
            else
            {
                if (($wrapper -ne $null) -and ($wrapper.Value -ne $null))
                {
                    $script:context.RetryQuery($wrapper.Value);
                    return;
                }
            }
        }
        catch [System.Net.WebException] {
            $response = $_.Exception.Response

            if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {

                $wrapper = [Microsoft.SharePoint.Client.ClientRequestWrapper]($_.Exception.Data["ClientRequest"]);
                $retry = $true


                $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
                $retryAfterInMs = $DefaultRetryAfterInMs;

                if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                    if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInterval)) {
                        $retryAfterInterval = $DefaultRetryAfterInMs;
                    }
                }
                else
                {
                    $retryAfterInterval = $backoffInterval;
                }

                Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInterval));
                #Add delay.
                Start-Sleep -m ($retryAfterInterval * 1000)
                #Add to retry count.
                $retryAttempts++;
                $backoffInterval = $backoffInterval * 2;
            }
            else {
                throw;
            }
        }
    }

    throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}

# Helper class
Add-Type -Language CSharp -ReferencedAssemblies $assemblies -TypeDefinition "
using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
public static class EnumOD4BListsHelper
{
    public static IEnumerable<SP.List> LoadListCollection(ClientContext context, Web web)
    {
        IEnumerable<SP.List> listCol = context.LoadQuery(
            web.Lists.Include(
                list => list.Title,
                list => list.RootFolder.ServerRelativeUrl
            ).Where(
                list => list.BaseTemplate == 700
            ));
        return listCol;
    }
}";

# Input password via console.
$password = Read-Host -Prompt "Please enter your password" -AsSecureString;

# Get ClientContext.
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password);
$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl);
$script:context.Credentials = $creds;

# Set UserAgent.
$script:context.add_ExecutingWebRequest({
    param ($source, $eventArgs);
    $request = $eventArgs.WebRequestExecutor.WebRequest;
    $request.UserAgent = "NONISV|Contoso|Application/1.0";
});

# Retrieve lists from OD4B site.
$web = $script:context.Web;
$listCol = [EnumOD4BListsHelper]::LoadListCollection($script:context, $web);
ExecuteQueryWithIncrementalRetry -retryCount 5;

# Found just 1 'Documents'.
if ($listCol.Count -eq 1)
{
    $dirs = $listCol.RootFolder.ServerRelativeUrl.Split("/");
    $nameOfDocuments = $dirs[-1];

    # When the MySiteDocumentLibrary is not 'Documents':
    if ($nameOfDocuments -ne "Documents")
    {
        Write-Output ("'MySiteDocumentLibrary(700)' is '{0}', not 'Documents'." -F $nameOfDocuments);
        if ($DiagOnly)
        {
            # diag only.
            Write-Output "Action required.";
        }
        else
        {
            # Fix the Url of 'Documents'
            $dirs[-1] = "Documents";
            $targetUrl = $dirs -join "/";
            $listCol.RootFolder.MoveTo($targetUrl);
            ExecuteQueryWithIncrementalRetry -retryCount 5;
            Write-Output ("Renamed '{0}' to 'Documents'." -F $nameOfDocuments);
        }
    }
    else
    {
        Write-Output "'MySiteDocumentLibrary(700)' is 'Documents'. No action required.";
    }
}
elseif ($listCol.Count -eq 0)
{
    throw "'MySiteDocumentLibrary(700)' does not exist.";
}
else
{
    throw "Two or more than 'MySiteDocumentLibrary(700)' exist.";
}