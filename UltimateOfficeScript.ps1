##Elevate Script=====================================================================================================================================================
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process==================================================================================================================================
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  
  exit
}
##Checks for existing Config.xml file in scripts directory.==========================================================================================================
$chkCFG = Test-Path -path '.\config.xml'
IF ($chkCFG = $false)
{
##Choose Client Ediditon, won't let you choose anything except 32 or 64 =============================================================================================
cls
DO  {
    $prodED = Read-Host "Input Office Client Edition, `"32`" or `"64`" bit"
    }Until ($prodED -eq 32 -or $prodED -eq 64)
##Choose Version of Office ==========================================================================================================================================
  DO {
  $prodID - $null 
  cls
    Write-Host "=====Which products and apps do you want to deploy?====="
    Write-Host "1: Office 2019 Pro Plus Retail" ##Office2019Retail
    Write-Host "2: Office 2019 Standard" ##ProPlus2019
    Write-Host "3: Office 2019 Volume" ##ProPlus2019Volume
    Write-Host "4: Office 2019 Home and Buisness" ##HomeBusiness2019Retail
    Write-Host "5: Office 365 Buisness Retail" ##O365BusinessRetail
    Write-Host "6: Office 365 Pro Plus Retail (E3, E4, E5)" ##Office365ProPlusRetail
    write-Host "7: Other"
    $prodIN = Read-Host "Select from Above"
     switch ($prodIN)
     {
             '1' {
                $prodID = "Office2019Retail"
                cls
           }  '2' {
                $prodID = "ProPlus2019"
                cls
           } '3' {
                $prodID = "ProPlus2019Volume"
                cls
           } '4' {
                $prodID = "HomeBusiness2019Retail"
                cls
           } '5' {
                $prodID = "O365BusinessRetail"
                cls
           } '6' {
                $prodID = "Office365ProPlusRetail"
                cls
           } '7' {
                  Do {
                  ## Other supported Product ID's for ODT config file ===============================================================================================
                    $prodOT = 0
                    cls
                    Write-Host "=====Press Enter to Return====="
                    Write-Host "1: Office 365 Business" ##O365BusinessRetail
                    Write-Host "2: Office 365 Small Business Premium" ##O365SmallBusPremRetail
                    Write-Host "3: Office Home and Student" ##HomeStudentRetail
                    Write-Host "4: Office Home and Student 2019" ##HomeStudent2019Retail
                    Write-Host "5: Office 365 Home Premimum" ##O365HomePremRetail
                    Write-Host "6: Office Personal 2019" ##Personal2019Retail
                    Write-Host "7: Office Professional 2019" ##Professional2019Retail
                    Write-Host "8: Project 2019 Pro" ##ProjectPro2019Retail
                    Write-Host "9: Project 2019 Pro Volume" ##ProjectPro2019Volume
                    Write-Host "10: Project Standard 2019" ##ProjectStd2019Retail
                    Write-Host "11: Project Standard 2019 Volume" ##ProjectStd2019Volume
                    Write-Host "12: Pubblisher 2019" ##Publisher2019Retail
                    Write-Host "13: Publisher 2019 Volume" ##Publisher2019Volume
                    Write-Host "14: Visio Professional 2019" ##VisioPro2019Retail
                    Write-Host "15: Visio Professional 2019Volume" ##VisioPro2019Volume
                    Write-Host "16: Visio Standard 2019" ##VisioStd2019Retail
                    Write-Host "17: Visio Standard 2019Volume " ##VisioStd2019Volume
                    Write-Host "18: Skype for Business Basic 2016" ##SkypeforBusinessEntryRetail
                    Write-Host "19: Skype for Business 2019" ##SkypeforBusinessRetail
                    Write-Host "20: Skype for Business 2019 Volume" ##SkypeforBusiness2019Volume
                    Write-Host "21: Skype for Business Basic 2019" ##SkypeforBusiness2019Retail
                    $prodOT = Read-Host "Select from Above"
                        switch ($prodOT)
                            {
                              '1' {
                             $prodID = "O365BusinessRetail"
                             cls
                           }  '2' {
                             $prodID = "O365SmallBusPremRetail"
                             cls
                           } '3' {
                             $prodID = "HomeStudentRetail"
                             cls
                           } '4' {
                             $prodID = "HomeStudent2019Retail"
                             cls
                           } '5' {
                             $prodID = "O365HomePremRetail"
                             cls
                           } '6' {
                             $prodID = "Personal2019Retail"
                             cls 
                          }  '7' {
                             $prodID = "Professional2019Retail"
                             cls
                          }  '8' {
                             $prodID = "ProjectPro2019Retail"
                             cls
                          } '9' {
                             $prodID = "ProjectPro2019Volume"
                             cls
                          } '10' {
                             $prodID = "ProjectStd2019Retail"
                             cls
                          } '11' {
                             $prodID = "ProjectStd2019Volume"
                             cls
                          } '12' {
                             $prodID = "Publisher2019Retail"
                             cls
                          }  '13' {
                             $prodID = "Publisher2019Volume"
                             cls
                           } '14' {
                             $prodID = "VisioPro2019Retail"
                             cls
                           } '15' {
                             $prodID = "VisioPro2019Volume"
                             cls
                           } '16' {
                             $prodID = "VisioStd2019Retail"
                             cls
                           } '17' {
                             $prodID = "VisioStd2019Volume"
                             cls
                           } '18' {
                             $prodID = "SkypeforBusinessEntryRetail"
                             cls 
                          }  '19' {
                             $prodID = "SkypeforBusinessRetail"
                             cls
                          }  '20' {
                             $prodID = "SkypeforBusiness2019Volume"
                             cls
                          } '21' {
                             $prodID = "SkypeforBusiness2019Retail"
                             cls
                          }
                        }
                      }UNTIL ($prodOT -lt 22)
           }
     }
   }UNTIL ($prodID -ne $null)
##Volume Licencing ==================================================================================================================================================
If ($prodID -like "*Volume")
        {
            $keyID = Read-Host "Copy and Paste Volume License Key"
            $pidKEY = " PID = `"$keyID`" "
        }
cls
##Building the Exclution List =======================================================================================================================================
DO { 
    Write-Host "========================Choose What Install Type=================="
    Write-Host "1: Basic 4 Applications (Outlook, Word, PP, Excel)"
    Write-Host "2: Basic 4 Applications Plus Teams"
    Write-Host "C: Custom"
    $prodDEX = Read-Host "Select from Above"
            $ac = "`n`t`t<ExcludeApp ID=`"Access`" />"
            $xl = "`n`t`t<ExcludeApp ID=`"Excel`" />"
            $gr = "`n`t`t<ExcludeApp ID=`"Groove`" />"
            $ly = "`n`t`t<ExcludeApp ID=`"Lync`" />"
            $od = "`n`t`t<ExcludeApp ID=`"OneDrive`" />"
            $on = "`n`t`t<ExcludeApp ID=`"OneNote`" />"
            $ol = "`n`t`t<ExcludeApp ID=`"Outlook`" />"
            $pp = "`n`t`t<ExcludeApp ID=`"PowerPoint`" />"
            $te = "`n`t`t<ExcludeApp ID=`"Teams`" />"
            $wo = "`n`t`t<ExcludeApp ID=`"Word`" />"
     switch ($prodDEX){
             '1' {
                $prodEX = -Join ($ac, $gr, $ly, $od, $on, $te)
           }  '2' {
                $prodEX = -Join ($prodEX, $gr, $ly, $od, $on)
           } 'c' {
           Do {
        ##Custom Exceptions =========================================================================================================================================
        cls
        Write-Host "=====Exclude======"
        Write-Host "1: Exclude Access"
        Write-Host "2: Exclude Excel"
        Write-Host "3: Exclude Groove"
        Write-Host "4: Exclude Lync"
        Write-Host "5: Exclude OneDrive"
        Write-Host "6: Exclude OneNote"
        Write-Host "7: Exclude Outlook"
        Write-Host "8: Exclude PowerPoint"
        Write-Host "9: Exclude Teams"
        Write-Host "10: Exclude Word"
        Write-Host "Q: Done with Exclusions"
            $prodCEX = Read-Host "Please Make Selection"
                switch ($prodCEX) {
               '1' {
                $ex = -Join ($ex, $ac)
             } '2' {
                $ex = -Join ($ex, $xl) 
             } '3' {
                $ex = -Join ($ex, $gr)
             } '4' {
                $ex = -Join ($ex, $ly)
             } '5' {
                $ex = -Join ($ex, $od)
             } '6' {
                $ex = -Join ($ex, $on)
             } '7' {
                $ex = -Join ($ex, $ol)
             } '8' {
                $ex = -Join ($ex, $pp)
             } '9' {
                $ex = -Join ($ex, $te)
             } '10' {
                $ex = -Join ($ex, $wo)
             }
                 }
                            }Until ($prodCEX -eq "Q")
                            cls
     }
     }
     }UNTIL ($prodDEX -lt 3 -OR $prodDEX -eq "Q")
     cls
Set-Content -Path "$PSScriptRoot\config.xml" -Value "<Configuration>`n`t<Add OfficeClientEdition=`"$OCE`" Channel=`"Monthly`">`n`t<Product ID=`"$prodID`"$pidKEY>`n`t<Language ID=`"en-us`" />$ex`n`t</Product>`n`t</Add>`n`t<Property Name=`"PinIconsToTaskbar`" Value=`"TRUE`" />`n`t<Updates Enabled=`"TRUE`" Channel=`"Monthly`" />`n`t<RemoveMSI />`n`t<Display Level=`"None`" AcceptEULA=`"TRUE`" />`n</Configuration>" -Force
}
##Installatoin======================================================================================================================================================
$chkBTS = Test-Path "$PSScriptRoot\Office"
switch ($chkBTS){
    'True' {
     Start-Process -filepath "setup.exe" -WorkingDirectory "$PSScriptRoot" -ArgumentList "/download config.xml" -Wait
    } 'False' {
     Start-Process -filepath "setup.exe" -WorkingDirectory "$PSScriptRoot" -ArgumentList "/download config.xml" -Wait
     Start-Process -filepath "setup.exe" -WorkingDirectory "$PSScriptRoot" -ArgumentList "/configure config.xml" -Wait
    }
}
