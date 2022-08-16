<#

.AUTHOR to YOU (Smart bunch + amazing trekkies)
    -   To help you, I spent time and effort to create this script and save you and your team some headaches 
    -   To help me, please give this script a "star", "Follow me", improve this script or share the link in social media 


.LINK
    -   This link:
	    https://github.com/ToGoWhereNoOne
    

.USE
    -   To query AD to get user's attributes -> BU Name, Group, User Name, Global ID, Email, User, Manager and AD Account Status


.USE CASES
    -   Internal or external auditor asks for this type of attributes on the fly.
    -   Supervisor needs details on users of specific solution/tool in your portafolio


.INSTRUCTIONS
    1) Pre-requisites to run this script:
        i.	Install the 'Remote Server Administration Tools'  
        ii.	If you are running this script from your company machine, connect to the VPN and stay connected
        iii.Input file with list of users. Must be in csv format
        iv.	Read the comments in this script to make a couple of changes 
    2) Start the PowerShell ISE console as administrator
    3) Copy this script into the console pane. Save it.
    4) Press F5 to run it
    5) It will prompt you to locate the .csv file
    6) In <2 minutes, the script will generate another .csv file on the desktop and open it up automatically for you
    7) Check the content for accuracy
    8) Save .csv file as an excel file to then add the functionality and/or formatting you desired

.NOTES
	Author: Q 
    1) If you are running the script at work or other sensitive environment, ensure your supervisor, the IT and/or Info Sec teams know about it for your own protection. These days there are tools that can detect PowerShell commands and can block you and trigger an alert and possibly an investigation.
    2) The script creator assumes no liability for the function, use or any consequences of this free script.
    3) This script was created in good faith to avoid doing tedious, redundant or time-consuming work.
    4) Get better than me:
    -   https://docs.microsoft.com/en-us/powershell/
    -   https://www.youtube.com/watch?v=at_MagcYK5M 
    -   https://www.youtube.com/watch?v=UVUd9_k9C6A
#> 

<#
The following block tells your machine to prompt you to select a .csv file
#> 

&{Function Get-FileName($ini)

    {[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename}

<#
The following block reads the data from the .csv file you selected, focuses on the 'email'column and checks if those emails are found in AD. 
If those emails are in AD, it asks AD for its corresponding attributes such as BU Name, Group, User Name, Global ID, Email, User, 
Manager and AD Account Status. Customized attributes such as "This is what you should do, PSO,'
'IAO (PSO) Action,' 'Completed by IAO (PSO)' were added to streamline overall audit process. 

Note: Under line 55 below replace [SOLUTION NAME] with whatever the name of the solution/tool is or whatever you want. 
#> 

    $path2Csv=Get-FileName
    $CsvUsers = Import-Csv $path2Csv
    $EmailHeader = ($CsvUsers | Get-Member | Where-Object { $_.Name -like '*email*' }).Name 
    $CsvUsersEmailsOnly = $CsvUsers.$EmailHeader 

    $DateFilename = Get-Date -format "yyyy.mm.dd"
    $Path2Desktop = "$($env:USERPROFILE)\Desktop\$DateFilename [SOLUTION NAME].csv" 

<#
The following block loads the emails one row at a time, finds the corresponding user in AD and saves its BU Name, Group, User Name, Global ID, Email, User, 
Manager and AD Account Status in an new .csv file on the desktop.
I formated the column names to headers that made sense to me but you can change those to whatever you like. You only have to change the text you see between the "N=" and ";" from row 68-76.
If you get error msgs, just undo what you did by pressing the undo button on the top menu. 
#> 

    Foreach($EachUserEmail in $CsvUsersEmailsOnly)
       { 
             Get-ADUser -SearchBase "OU=BDUsers,DC=BDX,DC=com" -Filter { UserPrincipalName -eq $EachUserEmail -OR mail -eq $EachUserEmail } `
             -Properties msExchExtensionAttribute21, msExchExtensionAttribute23, Name, SAMAccountName, mail, manager, Enabled  | 
             
             Select-Object @{N="BU Name";E={$_.msExchExtensionAttribute21}}, @{N="Group";E={$_.msExchExtensionAttribute23}}, @{N="User Name";E={$_.Name}}, 
                           @{N="Global ID";E={$_.SAMAccountName}}, @{N="Email";E={$_.UserPrincipalName }}, @{n='User Manager';e={(Get-ADUser $_.manager).name}}, 
                           @{N="AD Account Status";E={
                                If ($_.Enabled -eq $true) {"ENABLED Account"}
                                ElseIf ($_.Enabled -eq $false) { "DISABLED Account" }  }} |
                      Select-Object *, @{n=”This is what you should do, IAO : ”;e={ "Go to step 3 in job aid for instructions"}} |
                      Select-Object *, @{n=”IAO (Cybersecurity Officer) Action: ”;e={ "Maintain Access / Revoke Access"}} |
                      Select-Object *, @{n=”Completed by IAO (Cybersecurity Officer): ”;e={ "Not yet"}} |
                      Select-Object *, @{n=”Verified by MIAO (Cybersecurity Officer's Director): ”;e={ "Not yet"}} |

                     Export-Csv -LiteralPath $Path2Desktop -Append  -NoTypeInformation -Encoding UTF8 -Force | Wait-Job
         }

<#
The following block reads the data from the .csv file you selected, focuses on the 'email'column and checks that those emails are *not* found in AD. 
Then it appends those emails to the final .csv file along with corresponding group membership, user name and ID. This way, those accounts can be 
deleted in the solution you are helping audit.

Note: when working with other solutions, change "Teams" text below 
to whatever heading the other solution uses to refer to a BU or group. 
The same instruction goes for the rest of the attributes
#> 

     Foreach ($EachUsersInOriginalFile in $CsvUsers )
         {
            $userEmail = $EachUsersInOriginalFile."Email"
            $userBU = $EachUsersInOriginalFile."Teams"
            $userUserName = $EachUsersInOriginalFile."Name"
            $userGlobalID = $EachUsersInOriginalFile."ID"

            IF ((Get-ADUser -SearchBase "OU=BDUsers,DC=BDX,DC=com" -Filter { UserPrincipalName -eq $userEmail -OR mail -eq $userEmail} ).Count -eq 0 )  
                 {
                     $userEmail | select 'Teams', 'email', 'ID', 'Name' |
                      Select-Object @{N="BU Name";E={$userBU}},@{n="Group";e={ "-"}}, @{n="User Name";e={ $userUserName}}, @{n='User Manager';e={"-"}}, @{n="AD Account Status";e={ "DELETED Account"}}, @{N="Global ID";E={$userGlobalID}}, @{N="Email";E={$userEmail}} |
                      Select-Object *, @{n=”This is what you should do, IAO : ”;e={ "Go to step 3 in job aid for instructions"}} |
                      Select-Object *, @{n=”IAO (Cybersecurity Officer) Action: ”;e={ "Maintain Access / Revoke Access"}} |
                      Select-Object *, @{n=”Completed by IAO (Cybersecurity Officer): ”;e={ "Not yet"}} |
                      Select-Object *, @{n=”Verified by MIAO (Cybersecurity Officer's Director): ”;e={ "Not yet"}} |

                     Export-Csv -LiteralPath $Path2Desktop -Append  -NoTypeInformation -Encoding UTF8 -Force | Wait-Job
                 }
     
         }

                    # start Excel
                    Start-Process  "$Path2Desktop" -WindowStyle Maximized
}