Import-module ActiveDirectory

# Crude Random Password Generator 

Function GeneratePassword
{

	# How many Characters Minimum?

	$Length = 15

	# Create a password choosing everying from Character 34 to 127

	1..$Length | foreach { $Password += ([char]((Get-Random 93) + 34)) }

	# Convert to a Secure String

	$Password = Convertto-SecureString $Password -asplaintext -force

	Return $Password

}

Function GET-GroupInfo()
{

	Param(

		$City,

		$Division

	)

	$GroupName = $City.replace(" ", "") + "-" + $Division.replace(" ", "")

	$GroupDescription = "$Division in $City Access Group"

	# Return the Results (This is a feature new to version 3)

	[pscustomobject]@{Name = $Groupname; Description = $GroupDescription }

}

# Define OU at Base of AD for Offices

$BaseOU = "Offices"

# Provide List of Office Names

$CityOU = "Eindhoven", "Hapert", "Valkenswaard", "Weert", "Budel", "Tilburg"

# Provide List of Divisions Per Office

$DivisionOU = "Sales", "Administration", "IT", "Developers"

# DistinguishedName of Domain Root

$Domain = "DC=Ronsulting,DC=net"

# Whatever the name of the $BaseOU Combined with Domain

$CompanyPath = "OU=$BaseOU," + $Domain

# UPN Extension to the Domain

$UPN = "@ronsulting.net"

# Create BaseOU for Offices

NEW-ADOrganizationalUnit -name $BaseOU -path $Domain -ProtectedFromAccidentalDeletion:$False

# Create BaseOU for Management and servers

NEW-ADOrganizationalUnit -name "ServerManagement" -path $Domain -ProtectedFromAccidentalDeletion:$False
NEW-ADOrganizationalUnit -name "Servers" -path "OU=ServerManagement,$($Domain)"
NEW-ADOrganizationalUnit -name "Service Accounts" -path "OU=ServerManagement,$($Domain)"

#Redirect new computers to Servers OU
REDIRCMP "OU=Servers,OU=ServerManagement,$($Domain)"

# Create Server Managers Group
NEW-ADGroup -name "Manage SQL Servers" -GroupScope Global -Description "Manage All SQL Servers" -Path "OU=ServerManagement,$($Domain)"
NEW-ADGroup -name "Manage SQL Databases" -GroupScope Global -Description "Manage All SQL Databases on SQL Servers" -Path "OU=ServerManagement,$($Domain)"
NEW-ADGroup -name "Manage All Servers" -GroupScope Global -Description "Manage All Servers" -Path "OU=ServerManagement,$($Domain)"

# Gather through list of Cities

Foreach ($City in $CityOU) 
{

	# Create OU for City

	NEW-ADOrganizationalUnit -path $CompanyPath -name $City

	# Gather through list of Divisions

	Foreach ($Division in $DivisionOU)
 {

		# Create Division within City

		NEW-ADOrganizationalUnit -path "OU=$City,$CompanyPath" -name $Division

		# Create Group within Division and Description

		$Groupdata = GET-Groupinfo -City $City -Division $Division

		$GroupName = $Groupdata.Name

		$GroupDescription = $Groupdata.Description

		NEW-ADGroup -name $GroupName -GroupScope Global -Description $GroupDescription -Path "OU=$Division,OU=$City,$CompanyPath"

	}

}

# Pull together list of CSV raw data supplied from Generator

# 

$Names = IMPORT-CSV Scriptsamplenames.csv  -Encoding UTF8

# Generate 150 Random Users from pulled Raw data

#For ($x=0;$x -lt 150;$x++)
Foreach ($Name in $Names) {

 # Pick a Random First and Last Name

	#$Firstname=GET-Random $Names.Firstname
	#$Lastname=GET-Random $Names.Lastname

	#$Name = $Null
	#$Name = Get-Random $Names
	$Firstname = $Name.Firstname.trim()
	$Lastname = $Name.Lastname.trim()

 $Displayname = $Lastname + ", " + $Firstname

 

 # Pick a Random City

 $City = GET-RANDOM $Cityou

 # Pick a Random Division

 $Division = GET-RANDOM $DivisionOU

 $LoginID = $Firstname + $Lastname
 $LoginID = $LoginID.replace(" ", "")
 $LoginID = $LoginID.replace("'", "")
 $LoginID = $LoginID.replace(".", "")
 

 $UserPN = $LoginID + $UPN

 $Sam = $LoginID.padright(20).substring(0, 20).trim()

 # Define their path in Active Directory

 $ADPath = "OU=$Division,OU=$City,$CompanyPath"


 Write-Host "Creating User for $FirstName $Lastname with to division $Division"
 # Create the user in Active Directory

 #New-ADUser -GivenName $Givenname -Surname $Surname -DisplayName $Displayname -UserPrincipalName $UserPN -Division $Division -City $City -Path $ADPath -name $Displayname -SamAccountName $Sam -userpassword (GENERATEPASSWORD)
 New-ADUser -GivenName $Firstname -Surname $Lastname -DisplayName $Displayname -UserPrincipalName $UserPN -Division $Division -City $City -Path $ADPath -name $Displayname -SamAccountName $Sam -Accountpassword (Convertto-SecureString Welkom123 -asplaintext -force) -PasswordNeverExpires:$True -Enabled:$True

 # Add User to appropriate Security Group

 $Groupname = (GET-GroupInfo -city $City -division $Division).Name
 IF ($Division -eq "IT") {
		Write-Host "`tUser is IT Member, to add to server groups"
		ADD-ADGroupmember "Manage SQL Servers" -members $Sam
		ADD-ADGroupmember "Manage All Servers" -members $Sam
	}

 IF ($Division -eq "Developers") {
		Write-Host "`tUser is Developer, to add to Database Owner group"
		ADD-ADGroupmember "Manage SQL Databases" -members $Sam
	}

 ADD-ADGroupmember $Groupname -members $Sam

 # Enable the account for access

 #ENABLE-ADAccount $Sam
 #Set-ADUser -Identity $Sam -PasswordNeverExpires:$True

}

 
