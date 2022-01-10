# Import active directory module for running AD cmdlets
Import-Module activedirectory

#$ErrorActionPreference = "Stop"
#Set-StrictMode -Version "Latest"

#Assign the document folder
$files = Get-ChildItem "C:\Users\USER\Desktop\UserCreate\NewUserForms\"

forEach ($doc in $files) {

#Create word application object
$word = New-Object -ComObject Word.Application

#Assign document path
$doc = $doc.ToString()
$documentPath = "C:\Users\USER\Desktop\UserCreate\NewUserForms\$doc" 

#open the document
$document = $word.Documents.Open($documentPath)

#list all tables in doc
$document.Tables | ft


    $fname = $document.Tables[1].Cell(6,2).range.text
    $lname = $document.Tables[1].Cell(8,2).range.text
    $jobtitle = $document.Tables[1].Cell(15,2).range.text
    $department = $document.Tables[1].Cell(16,2).range.text
    $manager = $document.Tables[1].Cell(17,2).range.text
    $pagernumber = $document.Tables[1].Cell(4,2).range.text.Substring(17)

#get info from a certain part of the table in the word doc
               
$information = [PSCustomObject]@{
    fname = $fname -replace '\W',''
    lname = $lname -replace '\W',''
    jobtitle = $jobtitle
    department = $department -replace '\W',''
    manager = $manager -replace '\W',''
    pagernumber = $pagernumber -replace '\W',''
                }

Write-Output $fname
Write-Output $lname

#Write-Output $information
$information | Export-Csv -NoTypeInfo -Append -Force -Path C:\Users\USER\Desktop\UserCreate\newusers.csv

#Close the document
$document.close()

#Close Word
$word.quit()
                    }
                    
                    

#_________________________________EXCEL TIME_________________________________

  
#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv C:\Users\igallegos\Desktop\UserCreate\newusers.csv


#Loop through each row containing user details in the CSV file 
foreach ($User in $ADUsers)
{
	#Read user data from each field in each row and assign the data to a variable as below
	$fname = $User.fname
	$lname = $User.lname
    $fn1 = $fname.Substring(0,1)
    $username = "$fn1$lname"
    $username =  $username.ToLower()
    $email = "$fn1.$lname@polarahealth.com"
    $jobtitle = $User.jobtitle
    $department = $User.department 
    $pagernumber = $User.pagernumber
    $manager = $User.manager
    Write-Output $fname
    Write-Output $lname
    Write-Output $fn1
    Write-Output $username
    Write-Output $email

    #import the csv
    $Departments = Import-csv C:\Users\USER\Desktop\UserCreate\RUcodes.csv

    #loop through the CSV line by line
    foreach ($depnum in $Departments)
    {

        #assign
        $Number = $depnum.Code
    

        #assign department name
        $depname = $depnum.Description
    

	    if ($Number -eq $department)
        {
        $department = $depname
        }
    } 

    


	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $username})
	{
		 #If user does exist, give a warning
		 Write-Warning "A user account with username $username already exist in Active Directory."
           pause
	}
	else
	{
		#User does not exist then proceed to create the new user account 
       
		New-ADUser `
            -SamAccountName $username `
            -Name "$fname $lname" `
            -GivenName $fname `
            -Surname $lname `
            -Enabled $True `
            -DisplayName "$fname $lname" `
            -Company "Company Name" `
            -AccountPassword (convertto-securestring "password" -AsPlainText -Force)`
            -ScriptPath "Kix32.exe" `
            -HomeDirectory "U:\" `
            -Title "$jobtitle" `
            -Department "$department" `
            -OtherAttributes @{pager=$pagernumber} `
            -Description "$jobtitle" `           

            
            
            #-Manager "$manager" `
            


        #Set Remote Control Settings Permissions 
        $LdapUser = "LDAP://" + (Get-ADUser $username).distinguishedName
        $User = [ADSI] $LdapUser
        $User.InvokeSet("EnableRemoteControl",2)
        $User.InvokeSet("TerminalServicesHomeDirectory","U:")
        $User.setinfo()


        
        
     
        


        #set from template: memberships,office

        #set user options from template
        $template = Read-Host "Enter in the template for $username"
        $SAMusername = Get-ADUser -F {SamAccountName -eq $username}
        $memberships = Get-ADUser -F {SamAccountName -eq $template} -Properties memberof | Select-Object -ExpandProperty memberof
        ForEach ($Group in $memberships) {
        Add-ADPrincipalGroupMembership $username -MemberOf $memberships
                                     }
        

        $off = Get-ADUser -Identity $template -Properties Office | Select-Object -ExpandProperty Office 
        Set-ADUser $username -Office $off
        $man = Get-ADUser -Identity $template -Properties Manager | Select-Object -ExpandProperty Manager 
        Set-ADUser $username -Manager $man

        #UserLgon name set
        Set-ADUser $username –UserPrincipalName $username@company.org
        
        


    
        #Run the folder set up 

        #set V2 folder with full control 
        New-Item -Path "N:\" -Name "$username.V2" -ItemType "directory"
        $folderpath = "N:\$username.V2"
        $NewAcl = Get-Acl -Path "N:\$username.V2"
        $arguments = $username, "FullControl",”ContainerInherit,ObjectInherit”,"None", "Allow"
        $thingy = New-Object System.Security.AccessControl.FileSystemAccessRule $arguments
        $NewAcl.AddAccessRule($thingy)
        $NewAcl | Set-Acl $folderpath
        


        ##set user folder with full control 
        New-Item -Path "G:\" -Name $username -ItemType "directory"
        $folderpath = "G:\$username"
        $NewAcl = Get-Acl -Path "G:\$username"
        $arguments = $username, "FullControl",”ContainerInherit,ObjectInherit”,"None", "Allow"
        $isProtected = $true
        $preserveInheritance = $true
        $thingy = New-Object System.Security.AccessControl.FileSystemAccessRule $arguments
        $NewAcl.AddAccessRule($thingy)
        $NewAcl | Set-Acl $folderpath
        
        #disable inheritance
        $NewAcl.SetAccessRuleProtection($isProtected, $preserveInheritance)
        Set-Acl -Path "G:\$username" -AclObject $NewAcl
        

        #Remote Desktop Services User profile profile path set to "\\documents\profiles\$username"
        $User.invokeset("terminalservicesprofilepath","\\documents\profiles\$username") 
        $User.setinfo()
            
	}
}



#delete old csv file
Remove-Item -Path C:\Users\USER\Desktop\UserCreate\newusers.csv

#create new csv file with headers
    $props=[ordered]@{
     fname=''
     lname=''
     jobtitle=''
     department=''
     manager=''
     pagernumber=''
}
New-Object PsObject -Property $props | 
     Export-Csv C:\Users\USER\Desktop\UserCreate\newusers.csv -NoTypeInformation

pause
