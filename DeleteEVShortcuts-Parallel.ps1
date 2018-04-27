# Created for the purposes of running multiple instances to delete EV shortcuts using Powershell
# Author: Priom Mashuk
# V1.0

# Usage:
# Change $FilePath to the full UNC path of the input file
# Change $Username and $Password with a valid admin of the Exchange server
# Change the -ConnectionUri with the on-premises exchange server

# Variables
$incr = 10
$FilePath = "C:\Temp\input.csv"
$input = Get-Content $FilePath
$data = @("") * $input.Length
for ($i=0; $i -lt $input.length; $i++) {
   $data.setvalue($input[$i],$i)
}
$count = $input.Length
$marker = 0

#main code starts here
while ($count -ne 0) {
    #setup temp queue
    $temp = @("") * $incr
    for ($i = 0; $i -lt $incr; $i++) {
        $temp.SetValue($input[$marker],$i)
        $marker++
    }

    #start jobs
    for ($i = 0; $i -lt $temp.Length; $i++) {
        $jobName = "Job"+ " $i"
        Start-Job -Name $jobName -ScriptBlock {
            param(
                [string[]]$d,
                [int]$n
            )
            # Display which job is processing the task
            write-host "processing in job $n"

            # Exchange Initialisation
            #User details
            $Username = "USERNAME"
            $Password = ConvertTo-SecureString "PASSWORD" -AsPlainText -Force
            #Establish connection
            $userCredential = New-Object System.Management.Automation.PSCredential ($Username, $Password)
            # Change the connectionUri if not O365 to the on-premises location
            # To use kerberos connection - change HTTPS to HTTP and authentication parameter to kerberos
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://<SERVER FQDN>/PowerShell/ -Credential $userCredential -Authentication basic -AllowRedirection
            $Commands = Import-PSSession $Session -AllowClobber -DisableNameChecking

            get-mailboxfolderstatistics -identity $d[$n] | select legacydn
            #Search-Mailbox -Identity $d[$n] -SearchQuery "IPM.NOTE.EnterpriseVault.Shortcut" -GetContent
            Remove-PSSession $Session
        } -ArgumentList @($temp),$i #arguments to parse through to new instance
    }

    #wait for all jobs to finish before queing next task
    while (Get-Job -State "Running") { Start-Sleep 2 }
    
    #retrieve results of the job instances and output
    Get-Job | Receive-Job
    
    #manually remove all jobs created
    Remove-Job *

    #while loop management
    $count = $count - $incr
    if ($count -lt 0) {
        $count = 0
        break
    }
    #to add line spaces
    write-host "`n"
}
# end