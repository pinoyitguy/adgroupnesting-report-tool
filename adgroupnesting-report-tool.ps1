#region IMPORTS
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -AssemblyName System.Windows.Forms
#endregion

function Open-Form {
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace = [runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $newRunspace
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Name = "mainForm"
    $data = $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $psCmd = [powershell]::Create().AddScript({
        [xml]$xaml = @"
        <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="AD Security Group Nesting Report" Height="600" Width="600" MinHeight="550" MinWidth="600" WindowStartupLocation="CenterScreen">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="40"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition Height="30"/>
                    <RowDefinition />
                    <RowDefinition Height="25"/>
                    <RowDefinition Height="153"/>
                    <RowDefinition Height="45"/>
                </Grid.RowDefinitions>

                <!--UPN Controls-->
                <Label Grid.Column="0" Grid.Row="0" Content="User UPN:" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="20 0 0 0" />
                <TextBox Name="txtUser" Grid.Column="1" Grid.Row="0" Grid.ColumnSpan="3" Height="25" Margin="0 0 20 0" VerticalContentAlignment="Center" Background="#F0F0F0"/>

                <!--GDC Controls-->
                <Label Grid.Column="0" Grid.Row="1" Content="GDC Server:" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="20 0 0 0" />
                <TextBox Name="txtGdc" Grid.Column="1" Grid.Row="1" Height="25" VerticalContentAlignment="Center" Background="#F0F0F0"/>

                <!--Mode Controls-->
                <RadioButton Name="rbtMember" Grid.Column="2" Grid.Row="1" Content="Member" VerticalAlignment="Center" HorizontalAlignment="Right"/>
                <RadioButton Name="rbtMemberOf" Grid.Column="3" Grid.Row="1" Content="MemberOf" Margin="0 0 20 0" VerticalAlignment="Center" HorizontalAlignment="Right" />

                <!--XML Controls-->
                <Label Grid.Column="0" Grid.Row="2" Content="XML File Path:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="15 4 0 0" />
                <Grid Grid.Column="1" Grid.Row="2" Grid.ColumnSpan="3" Margin="0 5 20 0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="40"/>
                    </Grid.ColumnDefinitions>

                    <TextBox Name="txtXml" Grid.Column="0" Height="25" VerticalContentAlignment="Center" Background="#F0F0F0" IsReadOnly="True"/>
                    <Button Name="btnBrowse" Content="..." Grid.Column="1" Height="25" BorderThickness="0">
                        <Button.Style>
                            <Style TargetType="{x:Type Button}">
                                <Setter Property="Background" Value="#0958a3"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="{x:Type Button}">
                                            <Border Background="{TemplateBinding Background}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#627ea3"/>
                                    </Trigger>
                                    <Trigger Property="IsEnabled" Value="False">
                                        <Setter Property="Background" Value="#627ea3"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>

                <!--Main Reporting-->
                <TreeView Name="treeNest" Grid.Column="0" Grid.Row="3" Grid.ColumnSpan="4" Margin="20 7 20 0" Background="#F0F0F0"/>

                <!--Expand Controls-->
                <StackPanel Orientation="Horizontal" Grid.Column="2" Grid.Row="4" Grid.ColumnSpan="2" Margin="0 5 20 0" HorizontalAlignment="Right">
                    <Button Grid.Column="0" Name="btnExpand" Content="Expand" Width="60" Margin="0 0 10 0" VerticalAlignment="Top">
                        <Button.Style>
                            <Style TargetType="{x:Type Button}">
                                <Setter Property="Background" Value="#0958a3"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="{x:Type Button}">
                                            <Border Background="{TemplateBinding Background}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#627ea3"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>

                    <Button Grid.Column="1" Name="btnColl"  Content="Collapse" Width="60" VerticalAlignment="Top">
                        <Button.Style>
                            <Style TargetType="{x:Type Button}">
                                <Setter Property="Background" Value="#0958a3"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="{x:Type Button}">
                                            <Border Background="{TemplateBinding Background}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#627ea3"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </StackPanel>

                <!--Progress Controls-->
                <Label Grid.Column="0" Grid.Row="4" Content="Progress:" Margin="15 0 0 -10"/>
                <RichTextBox Name="rtbStat" Grid.Column="0" Grid.Row="5" Grid.ColumnSpan="4" Margin="20 0" Background="#F0F0F0" Block.LineHeight="1" VerticalScrollBarVisibility="Auto" IsReadOnly="True"/>
                <ProgressBar Name="prgStat" Grid.Column="0" Grid.Row="6" Grid.ColumnSpan="2" Foreground="#0958a3" VerticalAlignment="Top" Height="25" Margin="20 10 0 0"/>

                <!--Operation Buttons-->
                <Button Name="btnCancel" Grid.Column="2" Grid.Row="6" Content="Cancel" Height="25" Margin="10 10 10 0" VerticalAlignment="Top">
                    <Button.Style>
                        <Style TargetType="{x:Type Button}">
                            <Setter Property="Background" Value="#0958a3"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="{x:Type Button}">
                                        <Border Background="{TemplateBinding Background}">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="#627ea3"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
                <Button Name="btnRun" Grid.Column="3" Grid.Row="6" Content="Start" Height="25" Margin="0 10 20 0" VerticalAlignment="Top">
                    <Button.Style>
                        <Style TargetType="{x:Type Button}">
                            <Setter Property="Background" Value="#0958a3"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="{x:Type Button}">
                                        <Border Background="{TemplateBinding Background}">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="#627ea3"/>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter Property="Background" Value="#627ea3"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
            </Grid>
        </Window>
"@
        #region INITIALIZATION
        #---XAML parser---#
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
        $xaml.SelectNodes("//*[@Name]") | % {$syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}
        #endregion
        
        #region FUNCTIONS
        function Set-ControlDefaults {
            [System.Collections.ArrayList]$syncHash.TreeParents = @()
            $syncHash.TopFive = $null
            $syncHash.Enable = $true
            $syncHash.Indetermine = $false
            $syncHash.Status = $null
            $syncHash.Defaults = $false

            $syncHash.TreeParent = $null
            $syncHash.TreeTag = $null
            $syncHash.TreeParentTemp = $null
            $syncHash.TreeTagTemp = $null
            $syncHash.StatusTemp = $null
        }
        function Get-AdGroupHierarchy {
            $syncHash.stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $syncHash.Enable = $false
            $syncHash.Username = $syncHash.txtUser.Text
            $syncHash.GDC = "$($syncHash.txtGdc.Text):3268"
            $syncHash.treeNest.Items.Clear()
            $syncHash.rtbStat.Document.Blocks.Clear()
            $syncHash.TotalNestedGroupCount = 0
            $syncHash.XmlPath = $syncHash.txtXml.Text
            [System.Collections.ArrayList]$syncHash.GroupsDump = @()

            if($syncHash.rbtMember.IsChecked) {
                $syncHash.Mode = "Member"
            } else {
                $syncHash.Mode = "MemberOf"
            }

            $btnRunspace = [runspacefactory]::CreateRunspace()
            $btnRunspace.ApartmentState = "STA"
            $btnRunspace.ThreadOptions = "ReuseThread"
            $btnRunspace.Name = "btnRun"
            $btnRunspace.Open()
            $btnRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
            $cmdBtn = [powershell]::Create().AddScript({
                [System.Collections.ArrayList]$syncHash.Data = @()

                #region HELPERS
                function Update-Status([string]$Status) {
                    $syncHash.Status = $Status
                }
                function Get-FQDN {
                    [cmdletbinding()]
                    param([string]$DistinguishedName)
                    $domain = ($DistinguishedName -Split "," | ? {$_ -like "DC=*"} ) -join "." -replace ("DC=", "")
                    $domain
                }
                function Add-TreeItem {
                    Param (
                        $Name,
                        $Parent,
                        $Tag
                    )

                    $syncHash.TreeChild = $Name
                    $syncHash.TreeTag = $Tag
                    $syncHash.TreeParent = $Parent
                }
                function Test-TreeHierarchy([string]$Hierarchy) {
                    [array]$hArr = $Hierarchy -split ":" | Sort-Object
                    [array]$unique = $hArr | Get-Unique
        
                    if($unique.Count -lt $hArr.Count) {
                        $circular = $true
                    } else {
                        $circular = $false
                    }
        
                    return $circular
                }
                function Get-AdNestedGroup {
                    [CmdletBinding()]
                    Param (
                        [string]$Group,
                        [string]$TreeParent,
                        [System.Xml.XmlLinkedNode]$XmlParent,
                        [int]$Level = 1
                    )
                
                    $repeat = $false
                    if($Level -ge 2) {
                        $repeat = Test-TreeHierarchy -Hierarchy $TreeParent
                    }
                    
                    if(!$repeat) {
                        [void]$syncHash.Parents.Add($TreeParent)
                        try {
                            $srv = Get-FQDN -DistinguishedName $Group
                            $domainName = ($srv -split '\.')[0]
                            $parentGroup = Get-ADGroup $Group -Properties $syncHash.Mode -Server $syncHash.GDC -ErrorAction Stop
                            $parentName = $parentGroup.Name
                            [array]$parent = $parentGroup | Select-Object -ExpandProperty $syncHash.Mode | Sort-Object
                            if($parent.Count -gt 0) {
                                foreach($member in $parent) {
                                    try {
                                        $childGroup = Get-AdGroup -Identity $member -Properties $syncHash.Mode -Server $syncHash.GDC
                                        $groupName = $childGroup.Name

                                        if($childGroup.GroupCategory -eq "Distribution") {
                                            Continue
                                        }

                                        $parentCount = 0
                                        $childGroupArray = $childGroup | Select-Object -ExpandProperty $syncHash.Mode 
                                        foreach($child in $childGroupArray) {
                                            $chld = Get-AdGroup -Identity $child -Server $syncHash.GDC
                                            if($chld.GroupCategory -eq "Security") {
                                                $parentCount++
                                            }
                                        }
                                    } catch {
                                        Continue
                                    }
                                    
                                    $nextLvl = $Level + 1
                                    Update-Status "Adding $($groupName) to $parentName node at level $nextLvl..."
                                    Add-TreeItem -Name "[$parentCount] | [Level $nextLvl] | [$domainName] | $groupName" -Parent $TreeParent -Tag "$($TreeParent):$($groupName)"
                                    $table = [PSCustomObject][ordered]@{
                                        GroupName = $groupName
                                        ParentName = $parentName
                                        Level = $Level
                                    }
                                    $syncHash.TotalNestedGroupCount++
                                    [void]$syncHash.GroupsDump.Add($groupName)
                                    [void]$syncHash.Data.Add($table)
                                    $xmlName = ($groupName -replace ' ','_') -replace '&','and'
                                    $tmpXmlParentItem = $syncHash.xml.CreateNode("element",$xmlName,$null)
                                    $XmlParent.AppendChild($tmpXmlParentItem)

                                    if($parentCount -gt 0) {
                                        Get-AdNestedGroup -Group $member -TreeParent "$($TreeParent):$($groupName)" -XmlParent $tmpXmlParentItem -Level $nextLvl
                                    }
                                }
                            }
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("An error occurred in the function`n`n$($_.Exception)",'Error','OK','Error')
                        }
                    }
                }
                function Get-HighestGroupNesting([System.Collections.ArrayList]$Data) {
                    [System.Collections.ArrayList]$membCount = @()
                    $uniqueParents = $Data | Select-Object -ExpandProperty ParentName -Unique
                    foreach($parent in $uniqueParents) {
                        $parentData = $Data | Where-Object ParentName -eq $parent | Select-Object GroupName, ParentName, Level -Unique
                        $parentDataCount = ([array]($parentData | Where-Object Level -eq $parentData[0].Level)).Count

                        $d = [pscustomobject][ordered]@{
                            Parent = $parent
                            Count = $parentDataCount
                        }
                        [void]$membCount.Add($d)
                    }

                    return $membCount
                }
                #endregion

                #region MAIN
                try {
                    $user = Get-ADUser -Filter "UserPrincipalName -eq '$($syncHash.Username)'" -Server $syncHash.GDC -SearchBase ""
                    $xmlUser = ($user.Name -replace ' ','_') -replace '-','_'
                    #[array]$userGroups = Get-ADPrincipalGroupMembership $user.DistinguishedName -Server (Get-FQDN -DistinguishedName $user.DistinguishedName) -ErrorAction Stop | Where-Object GroupCategory -eq "Security"
                    [array]$userGroupsTemp = Get-ADUser -Identity $user.DistinguishedName -Server $syncHash.GDC -Properties MemberOf | Select-Object -ExpandProperty MemberOf
                    if($userGroupsTemp.Count -gt 0) {
                        $syncHash.Indetermine = $true
                        [xml]$syncHash.xml = New-Object System.Xml.XmlDocument
                        $dec = $syncHash.xml.CreateXmlDeclaration("1.0","UTF-8",$null)
                        $syncHash.xml.AppendChild($dec) | Out-Null
                        $root = $syncHash.xml.CreateNode("element",$xmlUser,$null)

                        $userGroups = @()
                        foreach($usrGrp in $userGroupsTemp) {
                            $groupInfo = Get-ADGroup -Identity $usrGrp -Server $syncHash.GDC
                            if($groupInfo.GroupCategory -eq "Security") {
                                $userGroups += $groupInfo.DistinguishedName
                            }
                        }

                        Update-Status "Adding root node..."
                        Add-TreeItem -Name "[$($userGroups.Count)] $($user.Name)" -Parent "root" -Tag "$($user.Name)"
                        Start-Sleep 1
                        foreach($userGroup in $userGroups) {
                            $userGroupDetails = Get-AdGroup -Identity $userGroup -Properties $syncHash.Mode -Server $syncHash.GDC
                            $userGroupName = $userGroupDetails.Name

                            Update-Status "Adding $userGroupName to root node..."
                            $syncHash.TotalNestedGroupCount++
                            [void]$syncHash.GroupsDump.Add($userGroupName)
                            $server = Get-FQDN -DistinguishedName $userGroup
                            $domainName = ($server -split '\.')[0]
                            #[array]$parentCountArr = Get-ADGroup -Identity $userGroup.DistinguishedName -Server $Global:gdc -Properties MemberOf | Select-Object -ExpandProperty MemberOf
                            [array]$parentCountArr = $userGroupDetails | Select-Object -ExpandProperty $syncHash.Mode
                            $parentCount = 0
                            foreach($child in $parentCountArr) {
                                try {
                                    $chld = Get-AdGroup -Identity $child -Server $syncHash.GDC
                                    if($chld.GroupCategory -eq "Security") {
                                        $parentCount++
                                    }
                                } catch {
                                    Continue
                                }
                            }
                            Add-TreeItem -Name "[$parentCount] | [Level 1] | [$domainName] | $userGroupName" -Parent "$($user.Name)" -Tag "$($user.Name):$($userGroupName)"
                            Start-Sleep 1
                            $table = [PSCustomObject][ordered]@{
                                GroupName = $userGroupName
                                ParentName = $user.Name
                                Level = 0
                            }
                            [void]$syncHash.Data.Add($table)
                            $xmlName = ($userGroupName -replace ' ','_') -replace '&','and'
                            $tmpXmlParent = $syncHash.xml.CreateNode("element",$xmlName,$null)
                            $root.AppendChild($tmpXmlParent)
                            [System.Collections.ArrayList]$syncHash.Parents = @()
                            Get-AdNestedGroup -Group $userGroup -TreeParent "$($user.Name):$($userGroupName)" -XmlParent $tmpXmlParent
                        }

                        #---XML---#
                        if($syncHash.XmlPath) {
                            $xmlFile = "$($syncHash.XmlPath)\$($user.Name)-NestedGroup-Report.xml"
                            $syncHash.xml.AppendChild($root) | Out-Null
                            $syncHash.xml.Save($xmlFile)
                        }
                        
                        #---Get Top 5---#
                        $topfive = Get-HighestGroupNesting -Data $syncHash.Data | Sort-Object Count -Descending | Select-Object -First 5
                        Update-Status "`nTop 5 Groups with Most Parents:"
                        foreach($top in $topfive) {
                            Start-Sleep -Milliseconds 500
                            Update-Status "$($top.Parent): $($top.Count)"
                        }

                        #---Total nested group count---#
                        Start-Sleep 1
                        Update-Status "`nTotal Number of Nested Group Memberships: $($syncHash.TotalNestedGroupCount)"

                        #---Number of duplicates count---#
                        Start-Sleep 1
                        [System.Collections.ArrayList]$tempArray = @()
                        $counter = 0
                        foreach($grp in $syncHash.GroupsDump) {
                            if($tempArray.Contains($grp)) {
                                $counter++
                            } else {
                                [void]$tempArray.Add($grp)
                            }
                        }
                        Update-Status "Total Number of Duplicate Group Memberships: $counter"

                        $syncHash.Stopwatch.Stop()
                        $hours = $syncHash.Stopwatch.Elapsed.Hours
                        $minutes = $syncHash.Stopwatch.Elapsed.Minutes
                        $seconds = $syncHash.Stopwatch.Elapsed.Seconds
                        $elapse = "Total script execution time is $hours hours $minutes minutes $seconds seconds"

                        if($syncHash.XmlPath) {
                            $success = "The process has been completed`n`nXML file saved to $xmlFile`n`n$elapse"
                        } else {
                            $success = "The process has been completed`n`n$elapse"
                        }
                        [System.Windows.Forms.MessageBox]::Show($success,'Success','OK','Information')
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("An error occurred on the main thread`n`n$($_.Exception)",'Error','OK','Error')
                }

                $syncHash.Defaults = $true
                if($syncHash.stopwatch.IsRunning) {
                    $syncHash.stopwatch.Stop()
                }
                #endregion
            })
            $cmdBtn.Runspace = $btnRunspace
            $cmdBtn.BeginInvoke() | Out-Null
        }
        function Expand-TreeViewItems {
            param (
                [System.Windows.Controls.TreeViewItem]$Item,
                [switch]$Shrink
            )

            if($PSBoundParameters["Shrink"]) {
                if($Item -ne $syncHash.treeNest.Items[0]) {
                    $Item.IsExpanded = $false
                }
                
                if($Item.Items.Count -gt 0) {
                    foreach($i in $Item.Items) {
                        Expand-TreeViewItems -Item $i -Shrink
                    }
                }
            } else {
                $Item.IsExpanded = $true
                if($Item.Items.Count -gt 0) {
                    foreach($i in $Item.Items) {
                        Expand-TreeViewItems -Item $i
                    }
                }
            }
        }
        #endregion

        #region CONTROLS UPDATER
        Set-ControlDefaults
        $syncHash.rbtMember.IsChecked = $true
        $updateControls = {
            if($syncHash.Defaults){
                Set-ControlDefaults
            }
            
            if($syncHash.TreeParent -ne $syncHash.TreeParentTemp -or $syncHash.TreeTag -ne $syncHash.TreeTagTemp) {
                if($syncHash.TreeParent -eq "root") {
                    $treeParent = $syncHash.treeNest
                } else {
                    $treeParent = $syncHash.TreeParents | Where-Object Hierarchy -eq $syncHash.TreeParent
                    $treeParent = $treeParent.Node
                }

                $tmpTreeChildItem = New-Object System.Windows.Controls.TreeViewItem
                $tmpTreeChildItem.Header = $syncHash.TreeChild
                $tmpTreeChildItem.Tag = $syncHash.TreeTag
                $tmpParent = [pscustomobject][ordered]@{
                    Hierarchy = $syncHash.TreeTag
                    Node = $tmpTreeChildItem
                }
                $syncHash.TreeParents.Add($tmpParent)
                [void]$treeParent.Items.Add($tmpTreeChildItem)
                $treeParent.IsExpanded = $true
                $syncHash.TreeParentTemp = $syncHash.TreeParent
                $syncHash.TreeTagTemp = $syncHash.TreeTag
            }

            if($syncHash.Status -ne $syncHash.StatusTemp) {
                $syncHash.rtbStat.AppendText("$($syncHash.Status)`n")
                $syncHash.rtbStat.ScrolltoEnd()
                $syncHash.StatusTemp = $syncHash.Status
            }
            
            $syncHash.prgStat.IsIndeterminate = $syncHash.Indetermine
            $syncHash.txtUser.IsEnabled = $syncHash.Enable
            $syncHash.txtGdc.IsEnabled = $syncHash.Enable
            $syncHash.btnRun.IsEnabled = $syncHash.Enable
            $syncHash.btnBrowse.IsEnabled = $syncHash.Enable
        }

        $syncHash.Window.Add_SourceInitialized({
            $timer = New-Object System.Windows.Threading.DispatcherTimer   
            $timer.Interval = [TimeSpan]"0:0:0.01"          
            $timer.Add_Tick($updateControls)            
            $timer.Start()                       
        })
        #endregion
        
        #region CONTROL EVENTS
        $syncHash.Window.Add_Closing({
            $rs = Get-Runspace -Name "btnRun"
            $iMain = Get-Runspace -Name "mainForm"

            if($syncHash.stopwatch.IsRunning -eq $true) {
                $closeWindow = [System.Windows.Forms.MessageBox]::Show("The operation is still running. Do you want to proceed?",'Exit Process','YesNo','Question')
                    if($closeWindow -eq "Yes") {
                        $rs.Close()
                        $rs.Dispose()
                        $iMain.CloseAsync()
                    } else {
                        $_.Cancel = $true
                    }
            } else {
                $rs.Close()
                $rs.Dispose()
                $iMain.CloseAsync()
            }
        })

        $syncHash.btnCancel.Add_Click({
            if($syncHash.Stopwatch.IsRunning) {
                $cancelProc = [System.Windows.Forms.MessageBox]::Show("The operation is still running. Do you want to proceed?",'Cancel Process','YesNo','Question')
                if($cancelProc -eq "Yes") {
                    $rs = Get-Runspace -Name "btnRun"
                    $rs.Close()
                    $rs.Dispose()
                    $syncHash.Defaults = $true
                    $syncHash.Stopwatch.Stop()
                }
            }
        })

        $syncHash.btnColl.Add_Click({
            foreach($item in $syncHash.treeNest.Items) {
                Expand-TreeViewItems -Item $item -Shrink
            }
        })

        $syncHash.btnExpand.Add_Click({
            foreach($item in $syncHash.treeNest.Items) {
                Expand-TreeViewItems -Item $item
            }
        })

        $syncHash.txtUser.Add_KeyDown({
            if($_.Key -eq "Enter") {
                if($syncHash.txtUser.Text -and $syncHash.txtGdc.Text) {
                    Get-AdGroupHierarchy
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Please provide UPN and GDC",'Start Process','OK','Warning')
                }
            }
        })

        $syncHash.txtGdc.Add_KeyDown({
            if($_.Key -eq "Enter") {
                if($syncHash.txtUser.Text -and $syncHash.txtGdc.Text) {
                    Get-AdGroupHierarchy
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Please provide UPN and GDC",'Start Process','OK','Warning')
                }
            }
        })

        $syncHash.btnRun.Add_Click({
            if($syncHash.txtUser.Text -and $syncHash.txtGdc.Text) {
                Get-AdGroupHierarchy
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please provide UPN and GDC",'Start Process','OK','Warning')
            }
        })

        $syncHash.btnBrowse.Add_Click({
            $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
            $foldername.Description = "Select a folder to save the XML report"
            if($syncHash.txtXml.Text){
                $foldername.SelectedPath = $syncHash.txtXml.Text
            } else {
                $foldername.RootFolder = "Desktop"
            }
            
            if($foldername.ShowDialog() -eq "OK"){
                $syncHash.txtXml.Text = $foldername.SelectedPath
            }
        })
        #endregion
        
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error += $Error
    })

    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()
    
    Return $syncHash
}

$prg = Open-Form

do {
    $main = Get-Runspace -Name "mainForm"
    Start-Sleep -Milliseconds 100
} while($main.RunspaceAvailability -ne "None")

$main.Dispose()