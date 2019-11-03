<#
    This Script is a gui to create an mst with default properties and the standard brand

    v1.0: 07/11/2018: Rory Pollock
          Current version GUID 63097938-e8b8-48a5-8e3b-6fc32e59d231
#>

$StartPath = $PSScriptRoot

Function New-MST {
    param(
        [Parameter(Mandatory=$true)]$MsiPath,
        #[Parameter(Mandatory=$true)][hashtable]$DefaultProperties,
        [string]$Clientname='Rory',
        $PackageName,
        [hashtable]$Properties,
        [hashtable]$SummaryInformation,
        [array]$RegAdditions
    )

    if(!$Properties){$Properties=@{}}
    <#
    $Properties = @{
        AgreeToLicense='Yes'
        REBOOT='ReallySuppress'
        ROOTDRIVE='C:\'
        ALLUSERS=1
        ARPNOMODIFY=1
        MSIRESTARTMANAGERCONTROL='Disable'
        ARPNOREPAIR=1
    }
    #>
    $Properties += @{CLIENTNAME=$Clientname}

    $BrandGuid = "868730B7-8980-4DB0-930D-130DBB3E828C"

    Write-Host "$MsiPath"
    $InicialItem = Get-Item $MsiPath

    if ($InicialItem.Extension -notmatch "\.msi"){
        Write-Host "MSI not found" -ForegroundColor Red
        "MSI not found"
    }
    else{
        $MSIObject = $InicialItem
    }

    Write-Host "Creating MST for $($MSIObject.name)" -ForegroundColor Cyan

    $CopyGuid = [guid]::NewGuid().Guid

    #$MSI = $MSIObject.FullName
    $MSI = "$env:TEMP\$CopyGuid.msi"
    
    Copy-Item -LiteralPath $MSIObject.FullName -Destination $MSI
    

    $msiname = $MSIObject.Name -replace "\.msi$",''
            
    if($PackageName){$MstName = $PackageName}
    else{$MstName = $msiname}

    $Properties += @{PACKAGENAME=$MstName}
    $MSTOutPath = "$($MSIObject.DirectoryName)\$MstName.mst"

    $N = 1
    while(Test-Path $MSTOutPath){
        $N++
        $MSTOutPath = "$($MSIObject.DirectoryName)\$MstName$("_")$N.mst"
    }

    Write-Host "OutputPath is $MSTOutPath" -ForegroundColor Green

    $msiOpenDatabaseModeReadOnly = 0
    $msiOpenDatabaseModeTransact = 1
    $msiTransformErrorNone = 0
    $msiTransformValidationNone = 0

    
    $guid = [guid]::NewGuid().Guid

    $ReferenceDatabasePath = $MSI
    $DiferenceDatabasePAth = "$env:TEMP\$guid.msi"

    Copy-Item -LiteralPath $ReferenceDatabasePath -Destination $DiferenceDatabasePAth -Force

    $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer         	 
    $ReferenceDatabase = $windowsInstaller.GetType().InvokeMember(
	    "OpenDatabase", 
	    "InvokeMethod", 
	    $Null, 
	    $windowsInstaller, 
	    @($ReferenceDatabasePath, $msiOpenDatabaseModeReadOnly)
    )  

    $DiferenceDatabase = $windowsInstaller.GetType().InvokeMember(
	    "OpenDatabase", 
	    "InvokeMethod", 
	    $Null, 
	    $windowsInstaller, 
	    @($DiferenceDatabasePAth, $msiOpenDatabaseModeTransact)
    ) 

    <#
    $tableExists =  $DiferenceDatabase.GetType().InvokeMember(
	    "TablePersistent",
	    "GetProperty",
	    $Null,
	    $DiferenceDatabase,
	    "Property"
    )	
    #>

    #Set Brand in regestry

    #if($Record){$PropertyString = "UPDATE Property SET Value = '$($Propertys.$Key)' WHERE Property = '$Key'"}
    #else{$PropertyString = "INSERT INTO ``Property`` (``Property``,``Value``) VALUES ('$Key','$($Propertys.$Key)')"}
    <#
    #>

    #Adds the feature
    $Feature = "INSERT INTO Feature (Feature,Feature_Parent,Title,Description,Display,Level,Directory_,Attributes) VALUES ('MST_registry_additions','','Brand_$BrandGuid','Registry additions','0','1','','0')"
    $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($Feature))
    $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
    $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

    $RegString1 = "INSERT INTO Registry (Registry,Root,``Key``,Name,Value,Component_) VALUES ('Brand1_$BrandGuid','2','SOFTWARE\[CLIENTNAME]\Install\msi\[PACKAGENAME]','Installed','1','Brand_$BrandGuid')"
    $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($RegString1))
    $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
        
    $RegString2 = "INSERT INTO Registry (Registry,Root,``Key``,Name,Value,Component_) VALUES ('Brand2_$BrandGuid','2','SOFTWARE\[CLIENTNAME]\Install\msi\[PACKAGENAME]','Time','[Time]','Brand_$BrandGuid')"
    $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($RegString2))
    $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null		
    $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
        
    $RegString3 = "INSERT INTO Registry (Registry,Root,``Key``,Name,Value,Component_) VALUES ('Brand3_$BrandGuid','2','SOFTWARE\[CLIENTNAME]\Install\msi\[PACKAGENAME]','Date','[Date]','Brand_$BrandGuid')"
    $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($RegString3))
    $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
    $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

    #$RegAdditions = 1..20 | %{New-Object -TypeName psobject -Property @{Registry="$(Get-Random)";Root="2";Key="SOFTWARE\Test\Bigtest_$_";Name="$_";Value="$_";Component="Brand_$BrandGuid"}}

    [array]$RegAdditions32 = $RegAdditions | ?{!$_.x64}
    [array]$RegAdditions64 = $RegAdditions | ?{$_.x64}

    if ($RegAdditions32){
        $RegAdditionsGUID = $([guid]::NewGuid().Guid).ToUpper()
        Foreach ($RegKey in $RegAdditions32){
            $RegString = "INSERT INTO Registry (Registry,Root,``Key``,Name,Value,Component_) VALUES ('$($RegKey.Registry)','$($RegKey.Root)','$($RegKey.Key)','$($RegKey.Name)','$($RegKey.Value)','REG_$RegAdditionsGUID')"
            $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($RegString))
            $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null		
            $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
        }
        
        $ComponentGuid = $([guid]::NewGuid().Guid).ToUpper() #Needs to be uppercase
        $Component = "INSERT INTO Component (Component,ComponentId,``Directory_``,Attributes,Condition,KeyPath) VALUES ('REG_$RegAdditionsGUID','{$ComponentGuid}','TARGETDIR','4','','$($RegAdditions32[0].Registry)')"
        $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($Component))
        $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
        $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

        #Adds the FeatureComponents
        $FeatureComponents = "INSERT INTO FeatureComponents (``Feature_``,``Component_``) VALUES ('MST_registry_additions','REG_$RegAdditionsGUID')"
        $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($FeatureComponents))
        $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
        $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
    }

    if ($RegAdditions64){
        $RegAdditionsGUID = $([guid]::NewGuid().Guid).ToUpper()
        Foreach ($RegKey in $RegAdditions64){
            $RegString = "INSERT INTO Registry (Registry,Root,``Key``,Name,Value,Component_) VALUES ('$($RegKey.Registry)','$($RegKey.Root)','$($RegKey.Key)','$($RegKey.Name)','$($RegKey.Value)','REG64_$RegAdditionsGUID')"
            $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($RegString))
            $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null		
            $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
        }
        
        $ComponentGuid = $([guid]::NewGuid().Guid).ToUpper() #Needs to be uppercase
        $Component = "INSERT INTO Component (Component,ComponentId,``Directory_``,Attributes,Condition,KeyPath) VALUES ('REG64_$RegAdditionsGUID','{$ComponentGuid}','TARGETDIR','260','','$($RegAdditions64[0].Registry)')"
        $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($Component))
        $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
        $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

        #Adds the FeatureComponents
        $FeatureComponents = "INSERT INTO FeatureComponents (``Feature_``,``Component_``) VALUES ('MST_registry_additions','REG64_$RegAdditionsGUID')"
        $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($FeatureComponents))
        $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
        $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
    }

    #Adds the component
    $ComponentGuid = $([guid]::NewGuid().Guid).ToUpper() #Needs to be uppercase
    $Component = "INSERT INTO Component (Component,ComponentId,``Directory_``,Attributes,Condition,KeyPath) VALUES ('Brand_$BrandGuid','{$ComponentGuid}','TARGETDIR','4','','Brand1_$BrandGuid')"
    $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($Component))
    $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
    $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

        
    #Adds the FeatureComponents
    $FeatureComponents = "INSERT INTO FeatureComponents (``Feature_``,``Component_``) VALUES ('MST_registry_additions','Brand_$BrandGuid')"
    $Insert = $DiferenceDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$Null,$DiferenceDatabase,($FeatureComponents))
    $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null	
    $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null

    <#
    $Query = "SELECT * FROM Feature WHERE ``Level`` = 1"
    $View = $DiferenceDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $DiferenceDatabase, ($Query))

    $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $Null)
    $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $Null, $View, $Null) #$null)
    $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $Null , $Record, 1)
    $Value | Write-Host -ForegroundColor Cyan
    $Value.GetType()
    #>

    #Set Properties
    foreach ($Key in $Properties.Keys){
        #Testif the Property exists already
        $Query = "SELECT Value FROM Property WHERE Property = '$($Key)'"
        $View = $DiferenceDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $DiferenceDatabase, ($Query))
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) | Out-Null
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($View) | Out-Null

        if($Record){$PropertyString = "UPDATE Property SET Value = '$($Properties.$Key)' WHERE Property = '$Key'"}
        else{$PropertyString = "INSERT INTO Property (Property,Value) VALUES ('$Key','$($Properties.$Key)')"}

        $Insert = $DiferenceDatabase.GetType().InvokeMember(
	        "OpenView",
	        "InvokeMethod",
	        $Null,
	        $DiferenceDatabase,
	        ($PropertyString)
        )
        $Insert.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $Insert, $Null) | Out-Null		
        $Insert.GetType().InvokeMember("Close", "InvokeMethod", $Null, $Insert, $Null) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Insert) | Out-Null
    }
    #<#
    #Commit all the changes
    $DiferenceDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $Null, $DiferenceDatabase, $Null) | Out-Null


    #Edit the summary information stream this needs to be done diferently
    $SummaryInformationLookup = @{
        'Codepage'=1
        'Title'=2
        'Subject'=3
        'Author'=4
        'Keywords'=5
        'Comments'=6
        'Template'=7
        'Last Saved By'=8
        'Revision Number'=9
        'Last Printed'=11
        'Create Time/Date'=12
        'Last Save Time/Date'=13
        'Page Count'=14
        'Word Count'=15
        'Character Count'=16
        'Creating Application'=18
        'Security'=19
    }

    #Sets the summary information table
    $SummaryInfo = $DiferenceDatabase.SummaryInformation(4)

    foreach ($key in $SummaryInformation.keys){
        if ($SummaryInformationLookup.Keys -contains $Key){
            if($SummaryInformation.$Key -match "^\[(.*)\]$"){
                try{
                    $Query = "SELECT Value FROM Property WHERE Property = '$($Matches[1])'"
                    $View = $DiferenceDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $DiferenceDatabase, ($Query))
                    $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) | Out-Null
                    $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                    $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $Null , $Record, 1)
                }
                catch{
                    $Value=$Matches[0]
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($View) | Out-Null
            }
            else{
                $Value=$SummaryInformation.$Key
            }
        
            $SummaryInfo.Property($($SummaryInformationLookup.$Key)) = $Value
        }
        else{
            Write-Host "$($SummaryInformation.keys) is not a key in the SummaryInformation table" -ForegroundColor Green 
        }
    }

    $SummaryInfo.Persist() | Out-Null
    $DiferenceDatabase.Commit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($SummaryInfo) | Out-Null

    #Generate a transform (the difference between our original MSI and our Backup MSI)	
    $transformSuccess = $DiferenceDatabase.GetType().InvokeMember(
	    "GenerateTransform", 
	    "InvokeMethod", 
	    $Null, 
	    $DiferenceDatabase, 
	    @($ReferenceDatabase,$MSTOutPath)
    )  

    #Create a Summary Information Stream for the MST
    $transformSummarySuccess = $DiferenceDatabase.GetType().InvokeMember(
	    "CreateTransformSummaryInfo", 
	    "InvokeMethod", 
	    $Null, 
	    $DiferenceDatabase, 
	    @($ReferenceDatabase,$MSTOutPath, $msiTransformErrorNone, $msiTransformValidationNone)
    )  
    #>
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($DiferenceDatabase) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ReferenceDatabase) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null

    Write-Host "Removing $DiferenceDatabasePAth" -ForegroundColor DarkCyan
    Remove-Item -LiteralPath $DiferenceDatabasePAth -Force #-ErrorAction SilentlyContinue 
    Write-Host "$env:TEMP\$CopyGuid.msi" -ForegroundColor DarkCyan
    Remove-Item -LiteralPath "$env:TEMP\$CopyGuid.msi" -Force
    Write-Host "Done" -ForegroundColor Magenta
    "$MstName Finished"
}

function Get-IniContent ($filePath){
    $ini = [ordered]@{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = [ordered]@{}
            $CommentCount = 0
            Continue
        }
        "^(;.*)$" # Comment
        {
            if (!($ini['Comments'])){$ini['Comments'] = [ordered]@{}}
            $name = $matches[1]
            $CommentCount = $CommentCount + 1
            $value = $section
            $ini['Comments'][$name] = $value
            Continue
        } 
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
            Continue
        }
    }
    return $ini
}

[xml]$xaml = @"
<Window 
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
x:Name="Window" Title="Rory's MST Maker" WindowStartupLocation = "CenterScreen"
SizeToContent = "WidthAndHeight" ShowInTaskbar = "True" Background = "White" ResizeMode="NoResize"> 
    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" > 
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Background = "White" Content = "Script Says:" Width = "100"/>
            <Label Height = "26" Name="MessageBox" Background = "White" Content = "Hello World"/>
        </StackPanel>
        <StackPanel  Name="MSIPathPanel" Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Width = "100" Background = "White" Content = "MSI Path:"/>
            <TextBox Name="MSIPathTB" Width='600' TextWrapping="Wrap" />
        </StackPanel>
        <StackPanel  Name="PackageNamePanel" Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Width = "100" Background = "White" Content = "Package Name:"/>
            <TextBox Name="PackageNamePathTB" Width='600' TextWrapping="Wrap" />
        </StackPanel>
        <StackPanel  Name="ClientPanel" Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Width = "100" Background = "White" Content = "Client:"/>
            <TextBox Name="ClientTB" Width='600' TextWrapping="Wrap" />
        </StackPanel>
        <Label Height = "26" Width = "100" Background = "White" Content = ""/>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight"> 
            <Label Height = "26" Width = "125" Background = "White" Content = "Property Table"/>
            <Button x:Name = "ResetPropertyTable" Height = "26" Content = 'Reset' ToolTip = "Add Property"  Width='172' />
        </StackPanel>
        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="PropertySP"> 
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight"> 
            <Label Height = "26" Width = "60" Background = "White" Content = ""/>
            <Button x:Name = "AddPropertyButton" Height = "26" Content = 'Add Property' ToolTip = "Add Property"  Width='237' />
        </StackPanel>
        <Label Height = "26" Width = "100" Background = "White" Content = ""/>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight"> 
            <Label Height = "26" Width = "125" Background = "White" Content = "Summary Information"/>
            <Button x:Name = "ResetSummaryInformation" Height = "26" Content = 'Reset' ToolTip = "Add Summary Information"  Width='172' />
        </StackPanel>
        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="SummaryInformationSP"> 
        </StackPanel>
                
        <Label Height = "26" Width = "100" Background = "White" Content = ""/>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight"> 
            <Label Height = "26" Width = "125" Background = "White" Content = "Registry Table"/>
            <Button x:Name = "ResetRegistryTable" Height = "26" Content = 'Reset' ToolTip = "Add Registry"  Width='172' />
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" Name="RegistryLabelSP"> 
            <Label Height = "26" Width = "155" Background = "White" Content = "Root"/>
            <Label Height = "26" Width = "400" Background = "White" Content = "Key"/>
            <Label Height = "26" Width = "100" Background = "White" Content = "Name"/>
            <Label Height = "26" Width = "100" Background = "White" Content = "Value"/>
            <Label Height = "26" Width = "60" Background = "White" Content = "x64"/>
        </StackPanel>
        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="RegistrySP"> 
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight"> 
            <Label Height = "26" Width = "60" Background = "White" Content = ""/>
            <Button x:Name = "AddRegistryButton" Height = "26" Content = 'Add Registry' ToolTip = "Add Registry"  Width='237' />
        </StackPanel>

        <Label Height = "26" Width = "100" Background = "White" Content = ""/>
        <Button x:Name = "CreateMSTButton" Height = "26" Content = 'Create MST' ToolTip = "Create MST" />

    </StackPanel>
</Window>
"@

$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

$CreateMSTButton = $Window.FindName('CreateMSTButton')

$MSIPathTB = $Window.FindName('MSIPathTB')
$PackageNamePathTB = $Window.FindName('PackageNamePathTB')
$ClientTB = $Window.FindName('ClientTB')
$PropertySP = $Window.FindName('PropertySP')

$ResetPropertyTable = $Window.FindName('ResetPropertyTable')

$MessageBox = $Window.FindName('MessageBox')
$AddPropertyButton = $Window.FindName('AddPropertyButton')

$ResetSummaryInformation = $Window.FindName('ResetSummaryInformation')
$SummaryInformationSP = $Window.FindName('SummaryInformationSP')

$ResetRegistryInformation = $Window.FindName('ResetRegistryTable')
$RegistryInformationSP = $Window.FindName('RegistrySP')
$AddRegistryButton = $Window.FindName('AddRegistryButton')


$ACTION = {
    $TrimmedPath = $MSIPathTB.Text  -replace '[\"]','' -replace "\'",''
    
    $PropertyTable = @{}
    foreach ($PropertyChild in $PropertySP.Children){
        if(($PropertyChild.Children[1].Text -and $PropertyChild.Children[3].Text) -and ($PropertyTable.keys -notcontains $PropertyChild.Children[1].Text)){
            $PropertyTable+=@{$PropertyChild.Children[1].Text = $PropertyChild.Children[3].Text}
        }
    }

    $SummaryInformation = @{}
    foreach ($SummaryInformationChild in $SummaryInformationSP.Children){
        if(($SummaryInformationChild.Children[1].Text -and $SummaryInformationChild.Children[3].Text) -and ($SummaryInformationChild.keys -notcontains $SummaryInformationChild.Children[1].Text)){}
        $SummaryInformation+=@{$SummaryInformationChild.Children[1].Text = $SummaryInformationChild.Children[3].Text}
    }

    $RegAdditions = 0
    $RegAdditionGUID = [guid]::NewGuid().Guid.toupper()
    $RegAditionCount = 0
    $RegistryAdditions = @()
    
    $RootLookup = @{
        "ALLUSERS dependent"="-1"
        "HKEY_CLASSES_ROOT"="0"
        "HKEY_CURRENT_USER"="1"
        "HKEY_LOCAL_MACHINE"="2"
        "HKEY_USERS"="3"
    }
    

    foreach ($RegistryInformationChild in $RegistryInformationSP.Children){

        if($RegistryInformationChild.Children[0].SelectedItem -and $RegistryInformationChild.Children[1].Text -and $RegistryInformationChild.Children[2].Text -and $RegistryInformationChild.Children[3].Text){
            $RegAditionCount++
            $NewReg = New-Object -TypeName psobject
            $NewReg | Add-Member -MemberType NoteProperty -Name Registry -Value "Reg$RegAditionCount$("_")$RegAdditionGUID"
            $NewReg | Add-Member -MemberType NoteProperty -Name Root -Value $RootLookup.$($RegistryInformationChild.Children[0].SelectedItem)
            $NewReg | Add-Member -MemberType NoteProperty -Name Key -Value $RegistryInformationChild.Children[1].Text
            $NewReg | Add-Member -MemberType NoteProperty -Name Name -Value $RegistryInformationChild.Children[2].Text
            $NewReg | Add-Member -MemberType NoteProperty -Name Value -Value $RegistryInformationChild.Children[3].Text
            $NewReg | Add-Member -MemberType NoteProperty -Name x64 -Value $RegistryInformationChild.Children[4].IsChecked

            $RegistryAdditions += $NewReg
        }
    }

    #$SummaryInformation = $(Get-IniContent -filePath "$StartPath\Settings.ini").SummaryInformation
    #$SummaryInformation = $(Get-IniContent -filePath "C:\Vms\Shared\Scripts\CreateMst\CreateMst\Settings.ini").SummaryInformation

    if($TrimmedPAth -and (Test-Path $TrimmedPAth)){
        $SPLAT = @{
            MsiPath=$TrimmedPath
            Properties=$PropertyTable
            SummaryInformation=$SummaryInformation
            RegAdditions=$RegistryAdditions
        }
        if($PackageNamePathTB.Text){$SPLAT+=@{PackageName=$PackageNamePathTB.Text}}
        if($ClientTB.Text){$SPLAT+=@{Clientname=$ClientTB.Text}}
        $GLOBAL:Message = New-MST @SPLAT
        $MessageBox.Content = $Message
    }
    else{
        $MessageBox.Content = "MSI not found please try again"
    }
}


$AddPropertyAction = {
    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'
    
    $Label1 = New-Object System.Windows.Controls.Label
    $Label1.Content = "Property:"
    $Label1.Width = 60
    $SPMother.AddChild($Label1)

    $TextBox1 = New-Object System.Windows.Controls.Textbox
    $TextBox1.Width = 237
    $TextBox1.Text = $P
    $SPMother.AddChild($TextBox1)

    $Label2 = New-Object System.Windows.Controls.Label
    $Label2.Content = "Value:"
    $Label2.Width = 50
    $SPMother.AddChild($Label2)

    $TextBox2 = New-Object System.Windows.Controls.Textbox
    $TextBox2.Width = 237
    $TextBox2.Text = $V
    $SPMother.AddChild($TextBox2)

    $RemoveButton = New-Object System.Windows.Controls.Button
    $RemoveButton.Height = 26
    $RemoveButton.Width = 26
    $RemoveButton.Content = '-'
    $RemoveButton.ToolTip = '-'
    $RemoveButton.Add_Click({
        $Parent = $this.Parent
        $GrandParent = $this.Parent.Parent
        $GrandParent.Children.Remove($Parent)
    })
    $SPMother.AddChild($RemoveButton)

    $PropertySP.AddChild($SPMother)
}

$AddSummaryInformationAction = {
    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'
    
    $Label1 = New-Object System.Windows.Controls.Label
    $Label1.Content = "Info:"
    $Label1.Width = 60
    $SPMother.AddChild($Label1)

    $TextBox1 = New-Object System.Windows.Controls.Textbox
    $TextBox1.Width = 237
    $TextBox1.Text = $P
    $TextBox1.IsReadOnly = $true
    $SPMother.AddChild($TextBox1)

    $Label2 = New-Object System.Windows.Controls.Label
    $Label2.Content = "Value:"
    $Label2.Width = 50
    $SPMother.AddChild($Label2)

    $TextBox2 = New-Object System.Windows.Controls.Textbox
    $TextBox2.Width = 237
    $TextBox2.Text = $V
    $SPMother.AddChild($TextBox2)

    $RemoveButton = New-Object System.Windows.Controls.Button
    $RemoveButton.Height = 26
    $RemoveButton.Width = 26
    $RemoveButton.Content = '-'
    $RemoveButton.ToolTip = '-'
    $RemoveButton.Add_Click({
        $Parent = $this.Parent
        $GrandParent = $this.Parent.Parent
        $GrandParent.Children.Remove($Parent)
    })
    $SPMother.AddChild($RemoveButton)

    $SummaryInformationSP.AddChild($SPMother)
}

#<Label Height = "26" Width = "40" Background = "White" Content = "Root"/>
#<Label Height = "26" Width = "400" Background = "White" Content = "Key"/>
#<Label Height = "26" Width = "150" Background = "White" Content = "Name"/>
#<Label Height = "26" Width = "150" Background = "White" Content = "Value"/>

$AddRegistryAction = {
    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'

    $TextBox1 = New-Object System.Windows.Controls.ComboBox
    $TextBox1.Width = 155
    $TextBox1.Items.Add("ALLUSERS dependent")
    $TextBox1.Items.Add("HKEY_CLASSES_ROOT")
    $TextBox1.Items.Add("HKEY_CURRENT_USER")
    $TextBox1.Items.Add("HKEY_LOCAL_MACHINE")
    $TextBox1.Items.Add("HKEY_USERS")
    # = "ALLUSERS dependant","HKCU","HKLM","HKU"
    $SPMother.AddChild($TextBox1)
    $TextBox2 = New-Object System.Windows.Controls.Textbox
    $TextBox2.Width = 400
    $SPMother.AddChild($TextBox2)

    $TextBox3 = New-Object System.Windows.Controls.Textbox
    $TextBox3.Width = 100
    $SPMother.AddChild($TextBox3)

    $TextBox4 = New-Object System.Windows.Controls.Textbox
    $TextBox4.Width = 100
    $SPMother.AddChild($TextBox4)

    $Checkbox = New-Object System.Windows.Controls.CheckBox
    #$Checkbox.DesiredSize.Width = 26
    $Checkbox.Margin = 5
    #$Checkbox.Width = 100
    $SPMother.AddChild($Checkbox)
    
    $RemoveButton = New-Object System.Windows.Controls.Button
    $RemoveButton.Height = 26
    $RemoveButton.Width = 26
    $RemoveButton.Content = '-'
    $RemoveButton.ToolTip = '-'
    $RemoveButton.Add_Click({
        $Parent = $this.Parent
        $GrandParent = $this.Parent.Parent
        $GrandParent.Children.Remove($Parent)
        if ($GrandParent.Children.count -lt 1){&$AddRegistryAction}
    })
    $SPMother.AddChild($RemoveButton)

    $RegistryInformationSP.AddChild($SPMother)
}

#default Property Section
$LoadDefault = {
    $SettingFile = Get-IniContent -filePath "$StartPath\Settings.ini"
    $DefaultProperties = $SettingFile.Property
    foreach($key in $DefaultProperties.Keys){
        $P = $key
        $V = $DefaultProperties.$key
        &$AddPropertyAction 
    }
    
}
#Default Summary table
$LoadDefaultS = {
    $SettingFile = Get-IniContent -filePath "$StartPath\Settings.ini"
    $SummaryInformation = $SettingFile.SummaryInformation
    foreach($key in $SummaryInformation.Keys){
        $P = $key
        $V = $SummaryInformation.$key
        &$AddSummaryInformationAction 
    }
}
&$LoadDefault
&$LoadDefaultS

$AddPropertyButton.add_click({
    $P = $null
    $V = $null
    &$AddPropertyAction 
})

$ResetPropertyTable.add_click({
    $count = $PropertySP.Children.Count
    if($count -gt 0){
        $range = $($count -1)..0 | %{$PropertySP.Children.RemoveAt($_)}
    }
    &$LoadDefault
})

$ResetSummaryInformation.add_click({
    $count = $SummaryInformationSP.Children.Count
    if($count -gt 0){
        $range = $($count -1)..0 | %{$SummaryInformationSP.Children.RemoveAt($_)}
    }
    &$LoadDefaultS
})

$AddRegistryButton.add_click({
    $P = $null
    $V = $null
    &$AddRegistryAction
})

$ResetRegistryInformation.add_click({
    $count = $RegistryInformationSP.Children.Count
    if($count -gt 0){
        $range = $($count -1)..0 | %{$RegistryInformationSP.Children.RemoveAt($_)}
    }
    &$AddRegistryAction
})

&$AddRegistryAction

$CreateMSTButton.add_click({&$ACTION})
$MSIPathTB.add_Keyup({if($_.Key -eq 'RETURN'){&$ACTION}})
$PackageNamePathTB.add_Keyup({if($_.Key -eq 'RETURN'){&$ACTION}})
$ClientTB.add_Keyup({if($_.Key -eq 'RETURN'){&$ACTION}})

$Window.Showdialog() | Out-Null
workflow ThisIsStupid {}


