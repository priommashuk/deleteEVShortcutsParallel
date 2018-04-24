# Created for the purposes of running multiple instances to delete EV shortcuts using Powershell
# Author: Priom Mashuk
# V0.5

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

# Exchange Initialisation
#User details
$Username = "USERNAME"
$Password = ConvertTo-SecureString "PASSWORD" -AsPlainText -Force

#Establish connection
$userCredential = New-Object System.Management.Automation.PSCredential ($Username, $Password)
Connect-MsolService -Credential $userCredential
# Change the connectionUri if not O365 to the on-premises location
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $userCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber


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
        #write-host $temp[$i]
        $jobName = "Job"+ " $i"
        Start-Job -Name $jobName -ScriptBlock {
            param(
                [string[]]$d,
                [int]$n
            )
            
            # Display which job is processing the task
            write-host "processing inside job $n"
            
            # Task to process
            # write-host $d[$n]   #debugging purposes only
            Search-Mailbox -Identity $d[$n] -SearchQuery "IPM.NOTE.EnterpriseVault.Shortcut" -DeleteContent

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