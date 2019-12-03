# Fix-OD4BDefaultDocuments
The default OneDrive for Business folder name is "Documents". If you change this name, you will encounter some error on OneDrive for Business site by this [article](https://support.office.com/en-us/article/renaming-the-documents-folder-in-file-explorer-breaks-navigation-to-onedrive-for-business-e3588fbb-1545-4f84-9309-233d10255291).
So, this script will help you to revert the default OneDrive for Business folder to "Documents".

## Prerequirements
- You need to download and install the latest SharePoint Online Client SDK by Nuget as well.
   - You need to access the following site. https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM
   - Download the nupkg.
   - Change the file extension to *.zip.
   - Unzip and extract those file.
   - Place "lib" folder under `C:\csom`.

## Usage
- Launch powershell.exe, and run as below:

`.\Fix-OD4BDefaultDocuments.ps1 -SiteUrl https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com -UserName john@contoso.onmicrosoft.com [-DiagOnly]`

## Parameter

- `-SiteUrl` Specifiy the URL of the OneDrive for Business site.
- `-UserName` Specify the user account who is site collection administrator of `-SiteUrl`.
- `-DiagOnly` (SwitchParameter) Use this parameter to only check the default folder name.
