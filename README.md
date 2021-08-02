# AzureDoc-Generator
 Azure Documentation Generator using Powershell

## Features

The program is work in progress. It generates DOCX file but needs more formatting

## Usage
You need to install DocumentFormat.OpenXML before using this program. 

### Installing DocumentFormat.OpenXML
Without Admin rights:

```powershell
Install-Package DocumentFormat.OpenXML -Scope CurrentUser
```
With Admin rights:

```powershell
Install-Package DocumentFormat.OpenXML
```
### Generating Document
Run the powershell script as follows, ResourceGroup being optional
```powershell
AzureDocumentation.ps1 -ResourceGroup "ResourceGroupName"
```
## Contributing
Pull requests are welcome.

## License
[MIT](https://choosealicense.com/licenses/mit/)
