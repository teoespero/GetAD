#################################################################################################
# 
# restore.ps1
# Teo Espero (#000891230)
# Cloud and Systems Administration (BS)
# Western Governors University
# Task 2
#          Objectives #1:
#          1. Create an Active Directory organizational unit (OU) named “finance.”
#          2. Import the financePersonnel.csv file (found in the “Requirements2” directory) 
#             into your Active Directory domain and directly into the finance OU. Be sure to 
#             include the following properties:
#                    
#                    - First Name
#                    - Last Name
#                    - Display Name (First Name + Last Name, including a space between)
#                    - Postal Code
#                    - Office Phone
#                    - Mobile Phone
#
#          Objectives #2:
#          1. Create a new database on the UCERTIFY3 SQL server instance called “ClientDB.”
#          2. Create a new table and name it “Client_A_Contacts.” Add this table to 
#             your new database.
#          3. Insert the data from the attached “NewClientData.csv” file (found in the 
#             “Requirements2” folder) into the table created in part B
#################################################################################################


#################################################################################################
## This function is created to accomplish objective #1

function objectiveOne(){

    $currPath = Get-Location
    $fileName = "financePersonnel.csv"


    if (Test-path $currPath"\"$fileName){
    
        ## Create an AD organizational unit (OU) "finance"

        New-ADOrganizationalUnit -Name finance -ProtectedFromAccidentalDeletion $false

        $ADPath = "OU=finance, DC=ucertify,DC=com"

        ## import the required CSV file

        $NewADUsers = Import-Csv $currPath"\"$fileName

        foreach ($ADUser in $NewADUsers){

            ## Define the vars to be used and initialize them with 
            ## the value from the CSV file

            $FirstName = $ADUser.First_Name
            $LastName = $ADUser.Last_Name
            $UserName = $ADUser.samAccount

            ## lookout for username that are more than the max for SamAccount

            if ($UserName.length -gt 20) {$UserName = $UserName.substring(0,20)}

            $City = $ADUser.City
            $County = $ADUser.County
            $ZipCode = $ADUser.PostalCode

            $OPhone = $ADUser.OfficePhone
            $MPhone = $ADUser.MobilePhone

            New-ADUser `
                -DisplayName $FirstName" "$LastName `
                -Name $UserName `
                -GivenName $FirstName `
                -Surname $LastName `
                -City $City `
                -StreetAddress $County `
                -PostalCode $ZipCode `
                -OfficePhone $OPhone `
                -MobilePhone $MPhone `
                -Path $ADPath

        }
    }
}

#################################################################################################
## This function is created to accomplish objective #2

function objectiveTwo(){

    ## find our path

    $currPath = Get-Location
    $sqlCodeFileName = "clients.sql"
    $csvFileName = "NewClientData.csv"

    ## Write-Host $currPath"\"$sqlCodeFileName
    ## Write-Host $currPath"\"$csvFileName

    

    ## load and register the SQL Server snap-ins and manageability assemblies

    Import-Module SQLPS -DisableNameChecking -Force

    ## create our object to connect the local SQL server

    ## get the servername

    $computerName = ".\UCERTIFY3"
    $serverName = New-Object -TypeName Microsoft.sqlserver.management.smo.server -ArgumentList $computerName
    $myDB = New-Object Microsoft.sqlserver.management.smo.database -ArgumentList $serverName, ClientDB
    $myDB.create()
    Invoke-Sqlcmd -ServerInstance $computerName -Database ClientDB -InputFile $currPath"\"$sqlCodeFileName

    ## table

    $myTable = 'dbo.Client_A_Contacts'
    $myDB = 'ClientDB'

    Import-Csv $currPath"\"$csvFileName | `
        ForEach-Object { Invoke-Sqlcmd -Database $myDB -ServerInstance .\UCERTIFY3 -Query `
            "insert into $myTable (`
                firstname,`
	            lastname,`
	            city,`
	            county,`
	            zip,`
	            officePhone,`
	            mobilePhone`
            )`
            values(`
                 '$($_.first_name)',`
                 '$($_.last_name)',`
                 '$($_.city)',`
                 '$($_.county)',`
                 '$($_.zip)',`
                 '$($_.officePhone)',`
                 '$($_.mobilePhone)'`
            )"
        }

}


#################################################################################################
## Main Program

try {
    $theCurrPath = Get-Location
    objectiveOne
    objectiveTwo
    Set-Location $theCurrPath
}
catch [System.OutOfMemoryException] {
    "An error occured that could not be resolved."
}
