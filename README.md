
.SYNOPSIS
This script checks and manages auto-reply settings for Exchange Online mailboxes.

.DESCRIPTION
The script performs the following tasks:
1. Installs the required module ExchangeOnlineManagement if not already installed.
2. Installs the NuGet provider if not already installed.
3. Sets the NuGet package source to the NuGet gallery.
4. Creates a directory for storing NuGet packages.
5. Downloads nuget.exe and adds the directory to the PATH.
6. Checks if the HtmlAgilityPack.dll is already installed. If not, installs it using NuGet.
7. Connects to Exchange Online using user-provided credentials.
8. Retrieves a list of mailboxes.
9. Checks if auto-reply is enabled for each mailbox.
10. Parses the auto-reply messages and removes the Byte Order Mark (BOM).
11. Generates a report in HTML format with the mailbox and auto-reply details.
12. Opens the generated report.
13. Provides an option to enable or disable auto-reply on a specific mailbox.

.PARAMETER None

.INPUTS
None

.OUTPUTS
None

.EXAMPLE
.\365_autoreplychecker.ps1
Runs the script and performs all the tasks.

.NOTES
- This script requires the ExchangeOnlineManagement module to be installed.
- The script prompts for Exchange Online credentials.
- The script requires the user to have the necessary permissions to manage Exchange Online mailboxes.
- The script generates an HTML report in the C:\tmp directory.
- The script requires an internet connection to download NuGet packages and the HtmlAgilityPack.dll.
