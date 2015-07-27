# Export format:
#    File Name
#    Policy Setting Name
#    Scope
#    Policy Path
#    Registry Information
#    Supported On
#    Help Text

# Get all ADMX  files from the current directory
$AdmxFiles = dir *.admx

$ns = @{ ns = "http://schemas.microsoft.com/GroupPolicy/2006/07/PolicyDefinitions"}
$nsDefined = [bool]1
$Delimiter = " ^ "
$CsvOutputFile = "ADMXSettingsData.csv"
$error_string = "String Not Found"

#generic function to display errors caught in various regions of the program
function Show-Error ($ADMX, $policy, $error_msg)
{
    $errorString = ""
    if ($ADMX) {$errorString += "$ADMX, "}
    if ($policy) {$errorString += "policy " + $policy.name+ ", "}
    if ($error_msg) {$errorString += $error_msg}
    Write-host -BackgroundColor Red "ERROR: " -NoNewline
    Write-Host  $errorString
}

#generic function to display warnings caught in various regions of the program
function Show-Warning ($ADMX, $policy, $error_msg)
{
    $errorString = ""
    if ($ADMX) {$errorString += "$ADMX, "}
    if ($policy) {$errorString += "policy " + $policy.name+ ", "}
    if ($error_msg) {$errorString += $error_msg}
    Write-host -BackgroundColor DarkYellow "Warning: " -NoNewline
    Write-Host  $errorString
}

Function Get-NameSpace ($ADMX, $NameSpaceRef)
{
   #Start by checking to see if a namespace is defined for this reference.If not, check to see if it is referring to itself and return "Namespace_Ref_Self"
   
   if($nsDefined){
      $NSref = $ADMX | Select-Xml -XPath "//ns:using[@prefix = '$NameSpaceRef']" -Namespace $ns
   }
   else{
      $NSref = $ADMX | Select-Xml -XPath "//using[@prefix = '$NameSpaceRef']"
   }

   if($nsDefined){

       if(($NSref.length -eq 0) -and (($ADMX | Select-Xml -XPath "//ns:target[@prefix = '$NameSpaceRef']" -Namespace $ns).count -ne 0))
       {
          return "Namespace_Ref_Self"
       }
   }
   else{
       if(($NSref.length -eq 0) -and (($ADMX | Select-Xml -XPath "//target[@prefix = '$NameSpaceRef']").count -ne 0))
       {
          return "Namespace_Ref_Self"
       }
   }
   
   return $NSref.Node.namespace
}

#find the file a namespace is defined in. if it's not well-known, try looking in [name].admx. otherwise, throw error.
Function Get-NameSpaceFile($NameSpace)
{
   #Check Well known name spaces to speed things up.
   #For any namespaces that get added, the corresponding ADMX file defining that namespace should be added here.
   switch ($NameSpace)
   {
      #for example, mmc.admx defines the namespace Microsoft.Policies.ManagementConsole
      "Microsoft.Policies.ManagementConsole" {return "mmc"}
      #and windows.admx defines the policy Microsoft.Policies.Windows
      "Microsoft.Policies.Windows" {return "windows"}
      "Microsoft.Policies.WindowsLogon" {return "logon"}
      "Microsoft.Policies.Products" {return "windowsproducts"}
      "Microsoft.Policies.Server" {return "windowsserver"}
      "Microsoft.Policies.MMCSnapIns" {return "mmcsnapins"}
      "Microsoft.Policies.GroupPolicy" {return "grouppolicy"}
      "Microsoft.Policies.InternetExplorer" {return "inetres"}
      "Microsoft.Policies.TerminalServer" {return "terminalserver"}
      "Microsoft.Policies.WindowsBackup" {return "windowsbackup"}
      "Microsoft.Policies.Backup" {return "windowsbackup"}
      "Microsoft.Policies.WindowsErrorReporting" {return "errorreporting"}
      "Microsoft.Policies.NetworkConnectivityAssistant" {return "nca"}

   }

   #try getting file name from namespace name 
   $NameSpaceParts = $NameSpace.Split(".")
   $FileName = $NameSpaceParts.Item($NameSpaceParts.length-1)
   
   if(Test-Path -Path ".\$FileName.admx")
   {
      $FileXML = [xml](Get-Content "$FileName.admx")
      if($nsDefined){
         $NameSpaceBool = $FileXML | Select-Xml -XPath "//ns:target[@namespace = '$NameSpace']" -Namespace $ns
      }
      else{
         $NameSpaceBool = $FileXML | Select-Xml -XPath "//target[@namespace = '$NameSpace']" 
      }
   }
      
   if($NameSpaceBool.count -eq 1)
   {
      if ($FileName -eq "")
      {
        #Namespace referenced and not found, please fix by modifying switch statement in this function
        return $error_string
      }
      return $FileName
   }
   else
   {
      #Namespace referenced and not found, please fix by modifying switch statement in this function
      return $error_string
   }
}
#recursively loop through the paths until we find the category an ADMX setting is listed under, then piece together the full path
Function Get-ParentPath ($PolicyPathRaw, $ADMX, $ADML)
{

    if($PolicyPathRaw.Contains(":"))
    {
        $PathInfo = $PolicyPathRaw.Split(":")
        $NameSpace = Get-NameSpace $ADMX $PathInfo.Item(0)
        #Check to see if the namespace was returned for the current file.
        if($NameSpace -eq "Namespace_Ref_Self")
        {
            $RefADMX = $ADMX
            $RefADML = $ADML
            $RefString = $PathInfo.Item(1)
        }
        #If not, get the namespace file and associated XML content.
        else
        {

                $RefFileName = Get-NameSpaceFile $NameSpace
                if ($RefFileName -eq $error_string) 
                {
                    show-error  "" "" "Get-Parent Path: Could not find definition for $policyPathRaw in $namespace"
                    return $error_string
                }
                else
                {
                    $RefADMXFile = $RefFileName + ".admx"
                    $RefADMLFile = $RefFileName + ".adml"
                    $RefString = $PathInfo.Item(1) #.ToLower()
                    $RefADMX = [xml](get-content .\$RefADMXFile)
                    $RefADML = [xml](get-content .\En-US\$RefADMLFile)
                }
        }
    }
    else
    {
        $RefADMX = $ADMX
        $RefADML = $ADML
        $RefString = $PolicyPathRaw
    }

    if($nsDefined)
    {
        $RefNameNode = $RefADMX | Select-Xml -XPath "//ns:category[@name = '$RefString']" -Namespace $ns
    }
    else
    {
        $RefNameNode = $RefADMX | Select-Xml -XPath "//category[@name = '$RefString']"
    }
    
    if(!$RefNameNode)
    {
        #referrant ADMX can have a different namespace
        $RefNameNode = $RefADMX | Select-Xml "//ns:category[@name = '$RefString']" -Namespace $ns
    }

    if(!$RefNameNode)
    {
        #this is also a valid code path for Reasons:
        $RefNameNode = $RefADMX | Select-Xml -XPath "//category[@name = '$RefString']"
    }
    #It's possible the ADMX we are referring to has a slightly wrong namespace
    if (!$RefNameNode)
    {
        $refNS = @{ ns = $RefADMX.policyDefinitions.xmlns}
        if(!$RefNameNode)
        {
            $RefNameNode = $RefADMX | Select-Xml -XPath "//ns:category[@name = '$RefString']" -Namespace $refNS
        }

        if(!$RefNameNode)
        {
            $RefNameNode = $RefADMX | Select-Xml -XPath "//category[@name = '$RefString']" -Namespace $refNS
        }
        if(!$RefNameNode)
        {
            Show-Error "" "" "Get-Parent-Path Could not find $refString in $refADMXFile"
            return $refString
        }
    }

    $PathNameVarRAW = $RefNameNode.Node.displayName
    $PathName = Get-String $RefADML $PathNameVarRAW

      
    # Check to see if this Category has a parent
    if($nsDefined)
    {
        $ParentCategoryRef = $RefNameNode | Select-Xml -XPath "ns:parentCategory" -Namespace $ns
    }
    else
    {
        $ParentCategoryRef = $RefNameNode | Select-Xml -XPath "parentCategory"
    }
    if($ParentCategoryRef.count -gt 0)
    {
        $ParentCategoryRaw = $ParentCategoryRef.Node.ref
        $ParentPath = Get-ParentPath $ParentCategoryRaw $RefADMX $RefADML
        $ParentPathName = $ParentPath + "\"
    }

    Return "$ParentPathName$PathName"

}


Function Get-RegistryInfo($ChildNodes, $DefaultKey, $Hive)
{
   foreach($childNode in $ChildNodes)
   {
      switch ($ChildNode.Name)
      {
         "enabledValue" #disabledValue elements ignored: duplicate registry keys / values are created when read this
         {
            if($ChildNode.decimal.valueName.length -gt 0)
            {
               $ChildRegistryPath = Get-RegistryPath $ChildNode.decimal $ChildRegistryPath $DefaultKey $Hive
            }
         }
         "enabledList" #disabledList elements ignored: duplicate registry keys / values are created when read this
         {
            if($ChildNode.ChildNodes.count -gt 0)
            {
               foreach($SubNode in $ChildNode.childNodes)
               {
                  $ChildRegistryPath = Get-RegistryPath $SubNode $ChildRegistryPath $DefaultKey $Hive
               }
            }
         }
         "elements"
         {
            if($ChildNode.ChildNodes.count -gt 0)
            {
               foreach($SubNode in $ChildNode.childNodes)
               {
                  $ChildRegistryPath = Get-RegistryPath $SubNode $ChildRegistryPath $DefaultKey $Hive
               }
            }
         }
      }

   }

   return $ChildRegistryPath
 
}

Function Get-RegistryPath($Node, $FullRegistryPath, $DefaultRegKey, $Hive)
{
    if($FullRegistryPath.length -gt 0){$FullRegistryPath += ", "}

    if($Node.key.length -gt 0)
    {
        $FullRegistryPath += $Hive + $Node.key 
    }
    else
    {
        $FullRegistryPath += $Hive + $DefaultRegKey
    }

    if($Node.valueName.length -gt 0)
    {
        $FullRegistryPath += "!" + $Node.valueName
    }

   return get-CleanText $FullRegistryPath
}

#remove linebreaks and etc so the csv doesn't get screwed up
Function Get-CleanText($RawText)
{
   #it's possible to call this function with no raw text

   #eat up: new line, carriage return, form feed, horizontal tab,
   #vertical tabs, and COMMAS (because we are exporting to CSV)
   $cleanText = $RawText -replace "`t|`n|`r|`f|`v|,|",""

   return $CleanText
}

#goes to an ADML file to look up the real-text
Function Get-String($ADML, $RawText)
{
   # it's possible to call this function with no raw text because it is recursively used in some cases
   #if (!$RawText) {write-host "Get-String called with empty RawText";}
   
   $nsold = ""
   $ADMLns = @{ ns = $adml.policyDefinitionResources.xmlns}
   if (!$ADML) 
   {
        show-error "" "" "Get-String called with non-existant ADML while looking for $rawtext";
   }
   #sanitize the string and looks for the referenced variable in ADML
   if($RawText.ToLower().Contains("`$(string".ToLower()))
   {
        $Var = $RawText.Substring(9,$RawText.Length-10)
        

        if($nsDefined)
        {
           $Node = $ADML | Select-Xml -XPath "//ns:string[@id = '$Var']" -Namespace $ns
        }

        else
        {
           $Node = $ADML | Select-Xml -XPath "//string[@id = '$Var']"
        }

        if(!$node)
        {
            #in the weird but possible event that an admx has a namespace but NOT the adml:
            $Node = $ADML | Select-Xml -XPath "//string[@id = '$Var']" -Namespace $ns
        }

        if(!$node)
        {
            #this is also a valid code path for Reasons:
            $Node = $ADML | Select-Xml -XPath "//ns:string[@id = '$Var']" -Namespace $ns
        }
        #It's possible the ADMX and ADML have two seperate namespaces
        if (!$Node)
        {
            if($nsDefined)
            {
               $Node = $ADML | Select-Xml -XPath "//ns:string[@id = '$Var']" -Namespace $admlNS
            }

            if(!$node)
            {
                $Node = $ADML | Select-Xml -XPath "//string[@id = '$Var']" -Namespace $ADMLns
            }

            if(!$node)
            {
                $Node = $ADML | Select-Xml -XPath "//ns:string[@id = '$Var']" -Namespace $ADMLns
            }
        }
        if(!$Node)
        {
            #cut this XML into pieces, this is my last resort
            Show-Error "" "" "String not found: $Var not found in namespace $ns, ADML namespace is $ADMLNs"
            return $error_string
        }

        return Get-CleanText $Node.Node.InnerText
                   
    }
    else
    {
        return Get-CleanText $RawText
    }
}

#MAIN FUNCTION
#Loop through every ADMX file and the policies defined within to gather information for spreadsheet
#Hardcoded for EN-US
#Expects to be run in the same DIR as ADMX files
#Expects a .\EN-US folder with ADML files
if($AdmxFiles.Count -gt 0)
{
   #Initialize Csv file information with Column headers
   $CsvFileInfo = "File Name" + $Delimiter + "Policy Setting Name" + $Delimiter + "Scope" + $Delimiter + "Policy Path" + $Delimiter + "Registry Information" + $Delimiter + "Supported On" + $Delimiter + "Help Text`n"

   ForEach($AdmxFile in $AdmxFiles)
   {
      Write-host "Processing" $AdmxFile.Name "..."

      $ADMX = [xml](get-content $AdmxFile)
      $ADMLname = $AdmxFile.BaseName + ".adml"
      $ADML = [xml](Get-Content .\EN-US\$ADMLname)

      # Check to see if a namespace is defined in the ADMX file. If it is use it if not use do not use a name space.
      if($admx.policyDefinitions.xmlns.length -gt 0)
      {
         $ns = @{ ns = $admx.policyDefinitions.xmlns}
         $nsDefined = [bool]1
         $Policies = $ADMX | Select-Xml -XPath "//ns:policy" -Namespace $ns
      }
      else
      {
         $nsDefined = [bool]0
         $Policies = $ADMX | Select-Xml -XPath "//policy"
      }

      if($Policies.Count -gt 0)
      {
         foreach($PolicyNode in $Policies)
         {
            Try
            {
               $Policy = $PolicyNode.Node
            }
            catch
            {
               Show-Error $AdmxFile "" "Failed to extract policy"
            }
            
            #Get Policy Name
            try
            {
                $NameVarRaw = $Policy.displayName
                $PolicyName = Get-String $ADML $NameVarRaw
                if ($policyName -eq $error_string)
                {
                    show-error $AdmxFile $Policy "Failed to get Policy Display Name for $namevarraw"
                    $PolicyName = $NameVarRaw
                }
            }
            catch
            {
                $PolicyName = $NameVarRaw
                show-error $AdmxFile $Policy "Failed to get Policy Display Name"
            }

            # Get Scope
            try
            {
                $Scope = Get-CleanText $Policy.class
            }
            catch
            {
                $Scope = ""
                show-error $AdmxFile $Policy "Failed to get scope of policy"
            }

            # Get Policy Path
            try
            {
                $PolicyPathRaw = $Policy.parentCategory.ref
                $PathName = Get-ParentPath $PolicyPathRaw $ADMX $ADML
                if ($PathName -eq $error_string)
                {
                    $PathName = $PolicyPathRaw
                    show-error $ADMXFile $policy "failed to get Get Parent Path for $policypathraw"
                }
            }
            catch
            {
                $PathName = $PolicyPathRaw
                show-error $ADMXFile $policy "failed to get Get Parent Path for $policypathraw"
                
            }

            #Get Registry Info
            try
            {
                $ThisPolicyOnly = [xml]$Policy.OuterXml
                
                $KeyName = $ThisPolicyOnly.policy.key
                $RegistryPath = ""

                if($Scope.ToLower() -eq "user")
                {
                   $Hive = "HKCU\"
                }
                else
                {
                   $Hive = "HKLM\"
                }

                if(($ThisPolicyOnly.policy.valueName).length -gt 0)
                {
                   $RegistryPath = $Hive + $KeyName + "!" + $ThisPolicyOnly.policy.valueName
                }
                
                if($ThisPolicyOnly.policy.ChildNodes.count -gt 0) # Should be able to optimize this by only looping through children if valueName attribute exists
                {
                    $RegistryPathOfChildren = Get-RegistryInfo $ThisPolicyOnly.policy.ChildNodes $KeyName $Hive
                    if(($RegistryPathOfChildren.length -gt 0) -and ($RegistryPath.length -gt 0))
                    {
                        $RegistryPath = $RegistryPath + "; " + $RegistryPathOfChildren
                    }
                    elseif(($RegistryPathOfChildren.length -gt 0) -and ($RegistryPath.length -eq 0))
                    {
                       $RegistryPath = $RegistryPathOfChildren
                    }
                }
                
                if($RegistryPath.Length -eq 0)
                {
                   $RegistryPath = Get-CleanText $Hive + $KeyName
                }
                #encapsulate in quotes to avoid the CSV freaking out over the comma
                #$registryPath = '"' + $RegistryPath + '"'
            }
            catch
            {
                $RegistryPath = ""
                show-error $ADMXFile $policy "Failed to get registry path"
            }
 
            # Get Supported On info
            try
            {
                $SupportedOnRaw = $Policy.supportedOn.ref

                #some teams have bad macros in ADMX
                #if their ADMX triggers this code path, file a bug on them to correct it
                #in the meantime add it to this list so the spreadsheet is not affected

                #correct bad macros for Windows 8.1
                if ($SupportedOnRaw -eq "windows:SUPPORTED_WindowsBlue") 
                {
                    $SupportedOnRaw = "windows:SUPPORTED_Windows_6_3"
                    Show-Warning $AdmxFile $Policy "Supported on: Used incorrect macro of supported_WindowsBlue instead of supported_Windows_6_3"
                }
                if ($SupportedOnRaw -eq "windows:SUPPORTED_Blue") 
                {
                    $SupportedOnRaw = "windows:SUPPORTED_Windows_6_3"
                    Show-Warning $AdmxFile $Policy "Supported on: Used incorrect macro of SUPPORTED_Blue instead of supported_Windows_6_3"
                }

                #correct bad macros for windows 8
                if ($SupportedOnRaw -eq "windows:SUPPORT_Win8") 
                {
                    $SupportedOnRaw = "windows:SUPPORTED_Windows8"
                    Show-Warning $AdmxFile $Policy "Supported on: Used incorrect macro of SUPPORT_Win8 instead of SUPPORTED_Windows8"
                }

                #correct bad macros for Windows 7
                if ($SupportedOnRaw -eq "windows:SUPPORTED_Win7") 
                {
                    $SupportedOnRaw = "windows:SUPPORTED_Windows7"
                    Show-Warning $AdmxFile $Policy "Supported on: Used incorrect macro of supported_Win7 instead of supported_Windows_7"
                }
                if ($SupportedOnRaw -eq "windows:SUPPORT_Win7") 
                {
                    $SupportedOnRaw = "windows:SUPPORTED_Windows7"
                    Show-Warning $AdmxFile $Policy "Supported on: Used incorrect macro of SUPPORT_Win7 instead of supported_Windows_7"
                }

                #look for the namespace specified before the colon
                if($SupportedOnRaw.Contains(":"))
                {
                   $SupportInfo = $SupportedOnRaw.Split(":")
                   $NameSpace = Get-NameSpace $ADMX $SupportInfo.Item(0)
                   #Check to see if the namespace was returned for the current file.
                   if($NameSpace -eq "Namespace_Ref_Self")
                   {
                       $RefADMX = $ADMX
                       $RefADML = $ADML
                       $RefString = $SupportInfo.Item(1)
                   }
                   #If not, get the namespace file and associated XML content.
                   else
                   {
                       $RefFileName = Get-NameSpaceFile $NameSpace
                       $RefADMXFile = $RefFileName + ".admx"
                       $RefADMLFile = $RefFileName + ".adml"
                       #write-host("refADMLfile = $refADMLFile")
                       $RefString = $SupportInfo.Item(1)
                       $RefADMX = [xml](get-content .\$RefADMXFile)
                       $RefADML = [xml](get-content .\En-US\$RefADMLFile)
                    }
                }
                else
                {
                   $RefADMX = $ADMX
                   $RefADML = $ADML
                   $RefString = $SupportedOnRaw
                }
                if($nsDefined){
                   $RefNameNode = $RefADMX | Select-Xml -XPath "//ns:* [@name = '$RefString']" -Namespace $ns
                }
                else{
                   $RefNameNode = $RefADMX | Select-Xml -XPath "//* [@name = '$RefString']"
                }
                #namespace may or may not actually be present, sanity check
                if (!$refNameNode)
                {
                    $RefNameNode = $RefADMX | select-xml -xpath "//* [@name = '$RefString']" -Namespace $ns
                }

                #if that fails again, it isn't a legtimate supportedOn and a bug needs to be filed against the owners of the ADMX file
                $SupportVarRAW = $RefNameNode.Node.displayName
                $SupportedOn = Get-String $RefADML $SupportVarRAW

            }
            catch
            {
                $error_msg = $_
                $SupportedOn = "Unknown"

                Show-Error $AdmxFile $Policy "SupportedOn: Could not find definition for $supportedOnRAW" + $error_msg
            }
            # Get Help Text 
            try
            {
                $HelpVarRaw = $Policy.explainText
                $HelpText = Get-String $ADML $HelpVarRaw

            }
            catch
            {
                $HelpText = ""
                Show-Error $AdmxFile $Policy "Help Text: Could not find definition for $HelpVarRaw"
            }

            #Save all of the information for the policy (File Name, Policy Setting Name, Scope, Policy Path, 
            # Registry Information, Supported On, Help Text).
            try
            {
                if($Scope.ToLower() -eq "both")
                {
                   $CsvFileInfo += $admxFile.Name + $Delimiter + $PolicyName + $Delimiter + "Machine" + $Delimiter + $PathName + $Delimiter + $RegistryPath + $Delimiter + $SupportedOn + $Delimiter + $HelpText + "`n"
                   #Duplicate all information and replace all instances of "HKLM\" with "HKCU\" 
                   $RegistryPath = $RegistryPath.Replace("HKLM\","HKCU\")
                   $CsvFileInfo += $admxFile.Name + $Delimiter + $PolicyName + $Delimiter + "User" + $Delimiter + $PathName + $Delimiter + $RegistryPath + $Delimiter + $SupportedOn + $Delimiter + $HelpText + "`n"                   
                }
                else
                {
                   $CsvFileInfo += $admxFile.Name + $Delimiter + $PolicyName + $Delimiter + $Scope + $Delimiter + $PathName + $Delimiter + $RegistryPath + $Delimiter + $SupportedOn + $Delimiter + $HelpText + "`n"
                }
            }
            catch
            {
                Show-Error $AdmxFile $Policy "Saving policy info to CSV format failed"
            }
         }
      }
      else
      {
         #some admx files only contain string definitions. this is not an error condition, but might be unexpected. 
         Show-Warning $AdmxFile "" "No policy settings found"
      }
   }

   try
   {
       Out-File -FilePath $CsvOutputFile -Encoding "default" -InputObject $CsvFileInfo
       Write-Host "Successfully output csv file."
   }
   catch
   {
        Show-Error "" "" "There was an error trying to output csv file."
   }
}
else
{
   Show-Error "" "" "No ADMX files found"
}

