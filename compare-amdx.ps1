#SUGGESTED USAGE: ./compare-admx.ps1 $old $new | out-file policies.txt
#SideIndicator can be used to find policy only in $new
#or policy only in $old
#Only get new policy: ./compare-admx.ps1 | ? {$_.sideIndicator -eq "=>"} | select-object inputobject
#only get old policy: ./compare-admx.ps1 | ? {$_.sideIndicator -eq "<="} | select-object inputobject

function compare-ADMX
{
    param( [Parameter(Position = 0, Mandatory=$true)]
            #Path of the older OS
           [string] $oldPath,
           [Parameter(Position = 1, Mandatory=$true)]
            #Path of the newer OS
           [string] $newPath,
           [switch] $onlyOld = $false
           
         )

    $oldADMX = gci $oldPath -force -filter "*.admx"
    $newADMX = gci $newPath -force -filter "*.admx"

    #get a list of all files that are in both old and new OS so we can look at any changes
    if (!$oldADMX) {write-host "Could not find .admx files in $oldPath"; return;}
    if (!$newADMX) {write-host "Could not find .admx files in $newPath"; return;}
    $both = diff $oldADMX $newADMX -ExcludeDifferent -IncludeEqual

    #and show the files that are included in one OS but not the other
    "=-=-=-=-=-=-=-=-=-=-=-=-=-=-"
    "ADMX files:"
    if ($onlyOld)
    {
        "Files and settings only present in $oldPath :"
        diff $oldADMX $newADMX | ? {$_.sideIndicator -eq "<="}
    }
    else
    {
        "ADMX files:"
        diff $oldADMX $newADMX 
    }
    "=-=-=-=-=-=-=-=-=-=-=-=-=-=-"
    "Settings:"
    #now go through the files in both OSes so we can compare content
    foreach ($file in $both)
        {
            $op = $oldPath + "\" + $file.InputObject.Name
            $np = $newPath + "\" + $file.InputObject.Name
            
            #if the onlyOld parameter is true, filter to sideIndicator <=
            if ($onlyOld)
            {
                $changes = diff $(Get-Content $op)  $(Get-Content $np) | ? {$_.sideIndicator -eq "<="} | select-object inputobject | fl
            }
            else 
            {
                $changes = diff $(Get-Content $op)  $(Get-Content $np) | fl
            }
            if ($changes)
            {
                #show filename
                $file.InputObject.Name
                #show diffed contents
                $changes
            }
        }

}