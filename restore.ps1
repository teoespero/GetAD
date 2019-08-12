#################################################################################################
# 


function task1(){

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

## Main Program

task1
