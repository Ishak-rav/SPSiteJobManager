# SPSiteJobManager

## Overview

This repository contains a PowerShell script designed to fetch and manage storage information of SharePoint sites. It's useful for administrators who need to monitor and manage SharePoint storage data efficiently.

## Files Description

- [Export-SharePointSitesList.ps1](Export-SharePointSitesList.ps1): Exports a list of SharePoint sites along with their storage details.
- [SPSiteJobManager.ps1](SPSiteJobManager.ps1): Manages jobs related to SharePoint sites, ensuring that storage information is up-to-date and accurate.
- [InstallModuleIfNeeded.ps1](InstallModuleIfNeeded.ps1): Installs the necessary modules if they are not already installed.
- [ImportModuleIfNeeded.ps1](ImportModuleIfNeeded.ps1): Checks whether the required PowerShell modules are installed and imports them if necessary.
- [ModuleChecker.ps1](ModuleChecker.ps1): Verifies the availability of specific PowerShell modules.
- [ProcessSite.ps1](ProcessSite.ps1): Contains functions for processing individual SharePoint sites.
- [GetCredentialWithRetries.ps1](ProcessSite.ps1): Aids in securely obtaining user credentials with retry logic.

## Getting Started

To use these scripts, you should have administrative access to the SharePoint sites you wish to manage. Ensure that you have the necessary permissions to execute scripts on your system.

### Prerequisites

- PowerShell 7
- PnP.PowerShell Module

### Installation

1. Clone the repository to your local machine.
2. Ensure that the execution policy for PowerShell scripts is appropriately set.
3. Install the SharePoint Online Management Shell if not already present.

### Usage

1. Navigate to the directory where the scripts are located.
2. Execute the scripts with the necessary parameters. For example:

```powershell
.\Export-SharePointSitesList.ps1
```

```powershell
.\SPSiteJobManager.ps1
```

## Contributing

Contributions are welcome! Please fork the repository, make your changes, and submit a pull request. We appreciate your efforts to improve the project.

## License

This project is licensed under the terms of the [LICENSE](LICENSE) file included in the repository.

## Acknowledgments

- Thanks to all the contributors who have invested their time in improving this script.
- Special thanks to the SharePoint community for the continuous support.

## Contact

For any queries or suggestions, please feel free to contact [Me](mailto:chennoufishak@gmail.com).
