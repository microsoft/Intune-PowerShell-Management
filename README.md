# Deprecation
This repo has been archived and will not receive further updates. For the latest version of the Intune PowerShell SDK, 
please use the landing page located at [https://aka.ms/IntuneScripts](https://aka.ms/IntuneScripts).

# Contributing
This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

# Legal Notices
Microsoft and any contributors grant you a license to the Microsoft documentation and other content
in this repository under the [Creative Commons Attribution 4.0 International Public License](https://creativecommons.org/licenses/by/4.0/legalcode),
see the [LICENSE](LICENSE) file, and grant you a license to any code in the repository under the [MIT License](https://opensource.org/licenses/MIT), see the
[LICENSE-CODE](LICENSE-CODE) file.

Microsoft, Windows, Microsoft Azure and/or other Microsoft products and services referenced in the documentation
may be either trademarks or registered trademarks of Microsoft in the United States and/or other countries.
The licenses for this project do not grant you rights to use any Microsoft names, logos, or trademarks.
Microsoft's general trademark guidelines can be found at http://go.microsoft.com/fwlink/?LinkID=254653.

Privacy information can be found at https://privacy.microsoft.com/en-us/

Microsoft and any contributors reserve all others rights, whether under their respective copyrights, patents,
or trademarks, whether by implication, estoppel or otherwise.

# Usage
## Setup
The Feature Modules and Samples depend on the [Intune PowerShell SDK](https://github.com/Microsoft/Intune-PowerShell-SDK).  Please install the Intune PowerShell SDK before running any of the code found in this repository.

## Tips and Tricks
 - Create TimeSpan objects using the [New-TimeSpan](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/new-timespan?view=powershell-6) cmdlet
 - Create DateTime or DateTimeOffset objects using the [Get-Date](https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Get-Date?view=powershell-6) cmdlet
 - If a parameter accepts an "Object" rather than a more specific type, use the documentation to identify what type of object it requires.  For example, if the documentation says that a parameter represents a property of type "microsoft.graph.mobileApp" or "microsoft.graph.deviceConfiguration", use the "New-MobileAppObject" or "New-DeviceConfigurationObject" cmdlets to create the respective objects.
