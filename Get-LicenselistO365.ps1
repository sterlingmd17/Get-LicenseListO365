#$mycreds = Get-Credential  #Eventually store this as hashed login.

$secpasswd = ConvertTo-SecureString "Testpass" -AsPlainText -Force

$mycreds = New-Object System.Management.Automation.PSCredential ("test@testemail.com", $secpasswd)

 

#Domain name to generate report

$DomainToFind = Read-Host -Prompt "What Domain?"

 

$Sku = @{

                "O365_BUSINESS_ESSENTIALS"                                      = "Office 365 Business Essentials"

                "O365_BUSINESS_PREMIUM"                                                         = "Office 365 Business Premium"

                "DESKLESSPACK"                                                                                                   = "Office 365 (Plan K1)"

                "DESKLESSWOFFPACK"                                                                      = "Office 365 (Plan K2)"

                "LITEPACK"                                                                                                              = "Office 365 (Plan P1)"

                "EXCHANGESTANDARD"                                                                     = "Exchange Online (Plan 1)"

                "STANDARDPACK"                                                                                                = "Enterprise Plan E1"

                "STANDARDWOFFPACK"                                                                                   = "Office 365 (Plan E2)"

                "ENTERPRISEPACK"                                                                                         = "Enterprise Plan E3"

                "ENTERPRISEPACKLRG"                                                                      = "Enterprise Plan E3"

                "ENTERPRISEWITHSCAL"                                                                               = "Enterprise Plan E4"

 

                "STANDARDPACK_STUDENT"                                                          = "Office 365 (Plan A1) for Students"

                "STANDARDWOFFPACKPACK_STUDENT"                                    = "Office 365 (Plan A2) for Students"

                "ENTERPRISEPACK_STUDENT"                                                    = "Office 365 (Plan A3) for Students"

                "ENTERPRISEWITHSCAL_STUDENT"                                         = "Office 365 (Plan A4) for Students"

                "STANDARDPACK_FACULTY"                                                            = "Office 365 (Plan A1) for Faculty"

                "STANDARDWOFFPACKPACK_FACULTY"                     = "Office 365 (Plan A2) for Faculty"

                "ENTERPRISEPACK_FACULTY"                                                     = "Office 365 (Plan A3) for Faculty"

                "ENTERPRISEWITHSCAL_FACULTY"                                          = "Office 365 (Plan A4) for Faculty"

                "ENTERPRISEPACK_B_PILOT"                                                      = "Office 365 (Enterprise Preview)"

                "STANDARD_B_PILOT"                                                                       = "Office 365 (Small Business Preview)"

               

                "VISIOCLIENT"                                                                                        = "Visio Pro Online"

                "POWER_BI_ADDON"                                                                                    = "Office 365 Power BI Addon"

                "POWER_BI_INDIVIDUAL_USE"                                      = "Power BI Individual User"

                "POWER_BI_STANDALONE"                                                             = "Power BI Stand Alone"

                "POWER_BI_STANDARD"                                                                                  = "Power-BI Standard"

                "PROJECTESSENTIALS"                                                                        = "Project Lite"

                "PROJECTCLIENT"                                                                                                 = "Project Professional"

                "PROJECTONLINE_PLAN_1"                                                              = "Project Online"

                "PROJECTONLINE_PLAN_2"                                                              = "Project Online and PRO"

                "ProjectPremium"                                                                                           = "Project Online Premium"

                "ECAL_SERVICES"                                                                                                 = "ECAL"

                "EMS"                                                                                                                        = "Enterprise Mobility Suite"

                "RIGHTSMANAGEMENT_ADHOC"                                                             = "Windows Azure Rights Management"

                "MCOMEETADV"                                                                                                             = "Microsoft Audio Conferencing"

                "SHAREPOINTSTORAGE"                                                                                    = "SharePoint storage"

                "PLANNERSTANDALONE"                                                                                  = "Planner Standalone"

                "CRMIUR"                                                                                                                           = "CMRIUR"

                "BI_AZURE_P1"                                                                                     = "Power BI Reporting and Analytics"

                "INTUNE_A"                                                                                                           = "Windows Intune Plan A"

                "PROJECTWORKMANAGEMENT"                                                                   = "Office 365 Planner Preview"

                "ATP_ENTERPRISE"                                                                                         = "Exchange Online Advanced Threat Protection"

                "EQUIVIO_ANALYTICS"                                                                       = "Office 365 Advanced eDiscovery"

                "AAD_BASIC"                                                                                                         = "Azure Active Directory Basic"

                "RMS_S_ENTERPRISE"                                                                        = "Azure Active Directory Rights Management"

                "AAD_PREMIUM"                                                                                                 = "Azure Active Directory Premium"

                "MFA_PREMIUM"                                                                                                = "Azure Multi-Factor Authentication"

                "STANDARDPACK_GOV"                                                                                    = "Microsoft Office 365 (Plan G1) for Government"

                "STANDARDWOFFPACK_GOV"                                                        = "Microsoft Office 365 (Plan G2) for Government"

                "ENTERPRISEPACK_GOV"                                                                             = "Microsoft Office 365 (Plan G3) for Government"

                "ENTERPRISEWITHSCAL_GOV"                                                   = "Microsoft Office 365 (Plan G4) for Government"

                "DESKLESSPACK_GOV"                                                                       = "Microsoft Office 365 (Plan K1) for Government"

                "ESKLESSWOFFPACK_GOV"                                                              = "Microsoft Office 365 (Plan K2) for Government"

                "EXCHANGESTANDARD_GOV"                                                        = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"

                "EXCHANGEENTERPRISE_GOV"                                                  = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"

                "SHAREPOINTDESKLESS_GOV"                                                   = "SharePoint Online Kiosk"

                "EXCHANGE_S_DESKLESS_GOV"                                     = "Exchange Kiosk"

                "RMS_S_ENTERPRISE_GOV"                                                            = "Windows Azure Active Directory Rights Management"

                "OFFICESUBSCRIPTION_GOV"                                                    = "Office ProPlus Government"

                "MCOSTANDARD_GOV"                                                                     = "Lync Plan 2G"

                "SHAREPOINTWAC_GOV"                                                                                 = "Office Online for Government"

                "SHAREPOINTENTERPRISE_GOV"                                                   = "SharePoint Plan 2G"

                "EXCHANGE_S_ENTERPRISE_GOV"                                               = "Exchange Plan 2G"

                "EXCHANGE_S_ARCHIVE_ADDON_GOV"                     = "Exchange Online Archiving"

                "EXCHANGE_S_DESKLESS"                                                                = "Exchange Online Kiosk"

                "SHAREPOINTDESKLESS"                                                                              = "SharePoint Online Kiosk"

                "SHAREPOINTWAC"                                                                                             = "Office Online"

                "YAMMER_ENTERPRISE"                                                                                   = "Yammer for the Starship Enterprise"

                "MCOLITE"                                                                                                               = "Lync Online (Plan 1)"

                "SHAREPOINTLITE"                                                                                          = "SharePoint Online (Plan 1)"

                "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"       = "Office ProPlus SMBiz"

                "EXCHANGE_S_STANDARD_MIDMARKET"                                 = "Exchange Online (Plan 1) Midmarket"

                "MCOSTANDARD_MIDMARKET"                                                     = "Lync Online (Plan 1)"

                "SHAREPOINTENTERPRISE_MIDMARKET"                              = "SharePoint Online (Plan 1)"

                "OFFICESUBSCRIPTION"                                                                = "Office ProPlus"

                "YAMMER_MIDSIZE"                                                                                      = "Yammer"

                "DYN365_ENTERPRISE_PLAN1"                                      = "Dynamics 365 Customer Engagement Plan Enterprise Edition"

                "ENTERPRISEPREMIUM_NOPSTNCONF"                     = "Enterprise E5 (without Audio Conferencing)"

                "ENTERPRISEPREMIUM"                                                                                    = "Enterprise E5 (with Audio Conferencing)"

                "MCOSTANDARD"                                                                                                = "Skype for Business Online Standalone Plan 2"

                "PROJECT_MADEIRA_PREVIEW_IW_SKU"                             = "Dynamics 365 for Financials for IWs"

                "STANDARDWOFFPACK_IW_STUDENT"                      = "Office 365 Education for Students"

                "STANDARDWOFFPACK_IW_FACULTY"                       = "Office 365 Education for Faculty"

                "EOP_ENTERPRISE_FACULTY"                                                     = "Exchange Online Protection for Faculty"

                "EXCHANGESTANDARD_STUDENT"                                               = "Exchange Online (Plan 1) for Students"

                "OFFICESUBSCRIPTION_STUDENT"                                           = "Office ProPlus Student Benefit"

                "STANDARDWOFFPACK_FACULTY"                                               = "Office 365 Education E1 for Faculty"

                "STANDARDWOFFPACK_STUDENT"                                              = "Microsoft Office 365 (Plan A2) for Students"

                "DYN365_FINANCIALS_BUSINESS_SKU"                 = "Dynamics 365 for Financials Business Edition"

                "DYN365_FINANCIALS_TEAM_MEMBERS_SKU" = "Dynamics 365 for Team Members Business Edition"

                "FLOW_FREE"                                                                                                         = "Microsoft Flow Free"

                "POWER_BI_PRO"                                                                                                = "Power BI Pro"

                "O365_BUSINESS"                                                                                                = "Office 365 Business"

                "DYN365_ENTERPRISE_SALES"                                        = "Dynamics Office 365 Enterprise Sales"

                "RIGHTSMANAGEMENT"                                                                                   = "Rights Management"

                "PROJECTPROFESSIONAL"                                                                 = "Project Professional"

                "VISIOONLINE_PLAN1"                                                                       = "Visio Online Plan 1"

                "EXCHANGEENTERPRISE"                                                                             = "Exchange Online Plan 2"

                "DYN365_ENTERPRISE_P1_IW"                                      = "Dynamics 365 P1 Trial for Information Workers"

                "DYN365_ENTERPRISE_TEAM_MEMBERS"                           = "Dynamics 365 For Team Members Enterprise Edition"

                "CRMSTANDARD"                                                                                                 = "Microsoft Dynamics CRM Online Professional"

                "EXCHANGEARCHIVE_ADDON"                                                       = "Exchange Online Archiving For Exchange Online"

                "EXCHANGEDESKLESS"                                                                       = "Exchange Online Kiosk"

                "SPZA_IW"                                                                                                              = "App Connect"

}

 

$SkuPrices = @{

"Office 365 Business Premium"              = 12.50

"Office 365 Business Essentials"           = 5.00

"Exchange Online (Plan 1)"                 = 4.00

"Enterprise Plan E3"                       = 20.00

"Office 365 Business"                      = 8.30

"Project Professional"                     = 30.00

"Office ProPlus"                           = 12.50

"Enterprise Plan E1"                       = 8.00

"Microsoft Audio Conferencing"             = 4.00

}

 

$Skuobject = ForEach($s in $sku.GetEnumerator()){[PSCustomObject]@{

    LicenseSku = $s.key

    LicenseFriendly = $s.Value

    Cost = if($skuprices[$s.value] -eq $null){ 0 } else{ $SkuPrices[$s.Value] }

    }

}

 

 

Connect-AzureAD -Credential $mycreds

Connect-MsolService -Credential $mycreds

 

#Store all domain names in array

$DomainNames = @()

 

#get all partner contracts in our portal

$Contracts = Get-MsolPartnerContract -DomainName $DomainToFind

 

#convert tenantid to domain names

foreach ($c in $Contracts)

{

$DomainNames += Get-MsolDomain -TenantId $c.TenantId | Where-Object {$_.Isinitial -eq $true}

}

 

#Grab tenant to be searched for and tenantID then store them in variables.

$DomainNames | Where-Object -Like "*$DomainToFind*" -Property "name" -OutVariable "TheDomain"

$TheDomainName = $TheDomain.name

$TheDomainNametitle = $TheDomainName.Split('.')[0]

$TheTenantID = (Invoke-WebRequest -Uri https://login.windows.net/$TheDomainName/.well-known/openid-configuration|ConvertFrom-Json).token_endpoint.Split(‘/’)[3]

 

 

#Find globally used licenses and available licenses for domain and convert them to friendly names.

$officelicense = (Get-MsolAccountSku -TenantId $TheTenantID | Select-Object -Property SkuPartNumber,ActiveUnits,ConsumedUnits | Sort-Object -Property skupartnumber, activeunits, consumedunits)

ForEach ($LicenseItem in $OfficeLicense){

 

   $SkuPartnumber = $LicenseItem.SkuPartNumber

   $LicenseItem.Skupartnumber = $Sku.$SkuPartnumber

 

}

 

#Find licenses applied to specific users and convert them to friendly names.

$UserObjects = Get-MsolUser -TenantId $TheTenantID

$UserObjects | Select-Object -Property UserprincipalName, Displayname, Licenses -OutVariable "UserObjectsFormatted"

ForEach ($a in $UserObjectsFormatted.licenses) {

$a.accountskuid = $a.accountskuid.split(':')[1]

}

#Clear-Variable -Name a

ForEach ($a in $UserObjectsFormatted.licenses){

$b = $a.AccountSkuId

$a.accountskuid = $sku.$b

$a = [system.String]::Join(" ", $a.AccountSkuid)

}

 

 

Function Get-DomainNames {

  $ar = @()

  ForEach ($user in $UserObjectsFormatted){

    $ar += $user.userPrincipalname.split("@")[1] } 

    $ar | Sort-Object -Unique

    

}

 

 

#loop through each user email and create variable for each new email and count how many licenses applied to that ending domain

 


$LicenseDomainTable =

    ForEach ($u in $UserObjectsFormatted){

    ForEach($lic in $u.Licenses.AccountSkuID){

    

    [PSCustomObject]@{

    DomainName = $u.userPrincipalname.split("@")[1]

    LicenseName = $lic

 

    }

    }

}

 

$LicenseCountTable =

   

    #somehow add a check too what domain the license belongs too.

 

    forEach ($i in Get-DomainNames ){

 

    ForEach($licname in $LicenseDomainTable | Where-Object -Property DomainName -Value $i -Like ) {

    $C = $LicenseDomainTable | Where-Object -Property DomainName -Value $i -Like | Select-Object -Property LicenseName | Where-Object -Property LicenseName -Value $licname.LicenseName -Like

    [PSCustomObject]@{

        

        Domain = $i

        License = $licname.LicenseName

        Count =  @($C).Count

 

        }

    }

   

}

$LicenseCountTable = $LicenseCountTable | Select-Object -Property License, Domain, Count -Unique

 

 

$body =

@"

<!DOCTYPE html>

<html>

  <head>

                <style>

                                table.blueTable {

                                border: 1px solid #1C6EA4;

                                background-color: #EEEEEE;

                                width: 75%;

                                text-align: left;

                                border-collapse: collapse;

                  }

                  table.blueTable td, table.blueTable th {

                                border: 1px solid #AAAAAA;

                                padding: 3px 2px ;

                  }

                  table.blueTable tbody td {

                                font-size: 13px;

                  }

                  table.blueTable tr:nth-child(even) {

                                background: #D0E4F5;

                  }

                  table.blueTable thead {

                                background: #1C6EA4;

                                background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);

                                background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);

                                background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);

                                border-bottom: 2px solid #444444;

                  }

                  table.blueTable thead th {

                                font-size: 15px;

                                font-weight: bold;

                                color: #FFFFFF;

                                border-left: 2px solid #D0E4F5;

                  }

                  table.blueTable thead th:first-child {

                                border-left: none;

                  }

                 

                  table.blueTable tfoot {

                                font-size: 14px;

                                font-weight: bold;

                                color: #FFFFFF;

                                background: #D0E4F5;

                                background: -moz-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);

                                background: -webkit-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);

                                background: linear-gradient(to bottom, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);

                                border-top: 2px solid #444444;

                  }

                  table.blueTable tfoot td {

                                font-size: 14px;

                  }

                  table.blueTable tfoot .links {

                                text-align: right;

                  }

                  table.blueTable tfoot .links a{

                                display: inline-block;

                                background: #1C6EA4;

                                color: #FFFFFF;

                                padding: 2px 8px;

                                border-radius: 5px;

                  }            </style>

    <meta charset="UTF-8">

    <title>title</title>

  <link rel="stylesheet" href="stylesheet.css" type="text/css">

  </head>

  <body>

 

<h3>All License info</h3>

  <table class="blueTable" style="margin-bottom: 50px">

                  <thead>

                                  <tr>

                                  <th>License</th>

                                  <th>Total Owned</th>

                                  <th>Available licenses</th>

                                  </tr>

                                </thead>

                                <tbody>

                                $(foreach($item in $officelicense){"

                                  <tr>

                                  <td>$($item.skupartnumber)</td>

                                  <td>$($item.activeunits)</td>

                                  <td>$($item.activeunits - $item.consumedunits)</td>

                                  </tr>

                                "})

                                </tbody>

  </table>

  <br></br>

 


 $(ForEach( $i in Get-DomainNames){"

<h3> $i </h3>

<table class='blueTable' style='margin-bottom: 50px'>

 

        <thead>

          <tr>

                                    <th>License</th>

            <th>Total applied</th>

            <th>Total Cost</th>

                                  </tr> 

                                </thead>

        

        

      <tbody>

      $($n = 0)

      $( ForEach($row in $licenseCountTable){"

      $($c = $Skuobject | Select-Object -Property LicenseFriendly, Cost -unique | Where-Object -Property LicenseFriendly -Value $row.License -Like | Select cost)

      $($TotalCostPerL = $c.cost * $row.count)

      $(if($row.Domain -eq $i){"

      $($TotalCost = $n += $TotalCostPerL )

      <tr>

       <td>$($row.License)</td>

                   <td>$($row.Count) </td>

       <td> $(("{0:C}" -f ($TotalCostPerL))) </td>

                  </tr>  "}) "})

      <tr>

           <td>All Licenses</td>

           <td></td>

           <td>$("{0:C}" -f ($n)) </td>   

             

      </tr>

 

                  </tbody>

 

  </table>

  <br></br> "})

  

 

  <table class="blueTable">

  <thead>

                <tr>

                <th>Email</th>

                <th>Display name</th>

                <th>License applied</th>

                </tr>

                </thead>

                <tbody>

               

  $(foreach($item in $userobjectsformatted){"

  <tr>

  <td>$($item.userprincipalname)</td>

  <td>$($item.displayname)</td>

  <td>$(if($item.licenses.accountskuid){$item.licenses.accountskuid}else{"No license applied"})</td>

  </tr>"})

  </tbody>

  </table>

  <br></br>

 

  </body>

</html>

 

"@

 

$body | Out-File $env:USERPROFILE\Desktop\$TheDomainNametitle.html

 

$PSEmailServer = 'Your email server or relay'

Send-MailMessage -from "from@test.com " -to "test@test.com" -Subject "License info update math of cost done and edited css for appearence What you think? " -BodyAsHtml $body

 

#add tag giving date license was assigned to end user as well as when taken away from user.

 #generate file for each domain.