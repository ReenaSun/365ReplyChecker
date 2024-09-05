function Start-Prereqs {
    # Import the required module
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
    }
    Import-Module ExchangeOnlineManagement

    # Install the NuGet provider if it's not already installed
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force
    }

    # Set the source to the NuGet gallery
    Set-PackageSource -Name NuGet -Location https://www.nuget.org/api/v2 -Trusted

    # Create the directory if it doesn't exist
    $nugetPath = "C:\nuget"
    if (!(Test-Path -Path $nugetPath)) {
        New-Item -ItemType Directory -Path $nugetPath
    }

    # Download nuget.exe
    $webClient = New-Object System.Net.WebClient
    $url = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
    $nugetExePath = Join-Path -Path $nugetPath -ChildPath "nuget.exe"
    $webClient.DownloadFile($url, $nugetExePath)

    # Add the directory to PATH
    $env:PATH += ";$nugetPath"

    # Define the path to the HtmlAgilityPack.dll
    $htmlAgilityPackPath = Resolve-Path "\\fileserver\users\reenas\scripps\packages\HtmlAgilityPack.1.11.64\HtmlAgilityPack.1.11.64\lib\netstandard2.0\HtmlAgilityPack.dll"

    # Check if HtmlAgilityPack is already installed
    if (!(Test-Path -Path $htmlAgilityPackPath)) {
        # Install HtmlAgilityPack if it's not already installed
        nuget install HtmlAgilityPack -OutputDirectory ./packages
    } else {
        # Load the HtmlAgilityPack assembly
        Add-Type -Path $htmlAgilityPackPath
    }
}

function Connect-Exchange {
    # Initialize a flag to indicate whether the credentials are valid
    $validCredentials = $false

    # Loop until valid credentials are provided
    while (-not $validCredentials) {
        # Ask for credentials
        $credential = Get-Credential -Message "Enter your Exchange Online credentials"

        # Try to connect with the provided credentials
        try {
            Write-Host "Connecting to Exchange Online..."
            Connect-ExchangeOnline -Credential $credential -ShowBanner:$false
            $validCredentials = $true
        }
        catch {
            Write-Warning "Invalid credentials. Please try again."
        }
    }
}

function Start-Report {
    #initialize response variable
    $skip = $null

    #skip to Set-AutoReplyState function
    while($skip -ne 'y' -and $skip -ne 'n') {
    $skip = Read-Host "Do you want to skip the report and go directly to enabling/disabling auto-reply on a specific mailbox? (Enter 'y' or 'n')"
    if ($skip -eq 'y') {
        return
    } elseif ($skip -eq 'n') {
        #continue
    } else {
        write-host "Invalid input. Please enter 'y' or 'n'"
    }
    }
        
        # Get the list of mailboxes
    Write-Host "Retrieving mailboxes..."
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object PrimarySmtpAddress

    # Initialize a list to hold the results
    $results = New-Object 'System.Collections.Generic.List[PSObject]'

    # Start the timer
    $timer = [Diagnostics.Stopwatch]::StartNew()

    # Loop through each mailbox
    foreach ($mailbox in $mailboxes) {
        Write-Host "Processing mailbox: $($mailbox.PrimarySmtpAddress)..."

        $autoreply = Get-MailboxAutoReplyConfiguration -Identity $mailbox.PrimarySmtpAddress

        # Check if autoreply is enabled
        if ($autoreply.AutoReplyState -eq 'Enabled' -or $autoreply.AutoReplyState -eq 'Scheduled') {
            # Get the auto-reply messages
            $externalMessage = $autoreply.ExternalMessage
            $internalMessage = $autoreply.InternalMessage

            # Remove BOM from the messages
            if ($null -ne $externalMessage) {
                $externalMessage = $externalMessage -replace '&#65279;', ''
            }
            if ($null -ne $internalMessage) {
                $internalMessage = $internalMessage -replace '&#65279;', ''
            }

            # Parse the HTML content
            $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument

            if ($null -ne $externalMessage) {
                $htmlDoc.LoadHtml($externalMessage)
                $spanElement = $htmlDoc.DocumentNode.SelectSingleNode("//span")
                $externalMessage = $spanElement.InnerText
            }
            if ($null -ne $internalMessage) {
                $htmlDoc.LoadHtml($internalMessage)
                $spanElement = $htmlDoc.DocumentNode.SelectSingleNode("//span")
                $internalMessage = $spanElement.InnerText
            }

            # Add the mailbox and autoreply details to the results
            $result = New-Object PSObject -Property @{
                Mailbox = $mailbox.PrimarySmtpAddress
                AutoReplyState = $autoreply.AutoReplyState
                ExternalMessage = $externalMessage
                InternalMessage = $internalMessage
                StartTime = $autoreply.StartTime
                EndTime = $autoreply.EndTime
            }

            $results.Add($result)
        }
    }

    # Stop the timer
    $timer.Stop()

    # Export the results to a HTML file
    $date = Get-Date -Format 'MMddyyyy-HHmm'
    $results | convertto-html -Property Mailbox, AutoReplyState, ExternalMessage, InternalMessage, StartTime, EndTime |Out-File -FilePath "C:\tmp\autoreply_results_$($date).html"

    # Display the total time taken
    Write-Host "Total time taken: $($timer.Elapsed.ToString())"

    # Open the HTML file
    $htmlFilePath = "C:\tmp\autoreply_results_$($date).html"
    Start-Process -FilePath $htmlFilePath
    } else {
        Write-Host "Invalid input. Please try again."
        Start-Report
    }
}

function Set-AutoReplyState {
    $consent = Read-Host "Do you want to enable or disable auto-reply on a specific mailbox? (Enter 'y' or 'n')"
    if ($consent -eq 'y') {
        $mailbox = Read-Host "Enter the email address of the mailbox"
        $enable = Read-Host "Do you want to enable auto-reply? (Enter 'y' or 'n')"

        if ($enable -eq 'y') {
        Set-MailboxAutoReplyConfiguration -Identity $mailbox -AutoReplyState Enabled
        } elseif ($enable -eq 'n') {
            Set-MailboxAutoReplyConfiguration -Identity $mailbox -AutoReplyState Disabled
        } else {
            Write-Host "Invalid input. Please try again."
        }
    } elseif ($consent -eq 'n') {
        Write-Host "Exiting..."
    } else {
        Write-Host "Invalid input. Please try again."
    }
}

# Main script
Start-Prereqs
Connect-Exchange
Start-Report
Set-AutoReplyState


# Disconnect the session
Disconnect-ExchangeOnline -Confirm:$false
