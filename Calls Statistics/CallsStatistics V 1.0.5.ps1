[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#--------------------------------------function----------------------------------------------#
function GenerateForm([string]$title, [int]$Width, [int]$Height){
 $form = New-Object System.Windows.Forms.Form
 $form.Text = $title
 $form.Width = $Width
 $form.Height = $Height
 $form.AutoSize = $true
 $form.StartPosition = "CenterScreen"
 $Icon = New-Object system.drawing.icon("\\docserver1\SYSTEM_DOCS\files\PowerShell\Icons\CallsStatistics.ico")
 $form.Icon = $Icon
return $form
}
function GenerateLabel([string]$text, [int]$x, [int]$y){
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $text
    $Label.Location  = New-Object System.Drawing.Point($x,$y)
    $Label.AutoSize = $true
    return $Label
}
function GenerateButton($window='', [string]$text, [int]$x, [int]$y, [scriptblock]$action=''){
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Location = New-Object System.Drawing.Point($x,$y)
    if($window){$window.Controls.Add($button)}
    if($action){$button.Add_Click( $action )}
    return $button
}
function GenGroupBox([string]$text, [int]$x, [int]$y, [string]$Range, [int]$width, [int]$height){
  $groupBox = New-Object System.Windows.Forms.GroupBox
  $groupBox.Controls.AddRange(@($Range))
  $groupBox.Location = New-Object System.Drawing.Point($x, $y)
  $groupBox.Size = New-Object System.Drawing.Size($width, $height)
  $groupBox.Text = $text
  return $groupBox
  }
function GenTextBox([int]$x, [int]$y, [int]$width, [int]$height, $text = ''){
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.AutoSize= $false
$textBox.Location = New-Object System.Drawing.Size($x,$y) 
$textBox.Size = New-Object System.Drawing.Size($width,$height)
$textBox.Text = $text
return $textBox
}
function GenComboBox([array]$data, [int]$x, [int]$y, [int]$width, [int]$height){
    $ComboBox = New-Object System.Windows.Forms.ComboBox
    $ComboBox.DataSource = @($data)
    $ComboBox.Location  = New-Object System.Drawing.Point($x,$y)
    $ComboBox.Size = New-Object System.Drawing.Size($width,$height)
    $ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList 
    return $ComboBox
    }
function GeneraterichText([string]$text, [int]$x, [int]$y, [int]$width, [int]$height){
    $richTextBox = New-Object System.Windows.Forms.RichTextBox
    $richTextBox.Text= $text
    $richTextBox.Location = New-Object System.Drawing.Point($x,$y)
    $richTextBox.Size = New-Object System.Drawing.Size($width, $height)
    $richTextBox.Font = New-Object System.Drawing.Font("'arial'",10,[System.Drawing.FontStyle]::Regular)
    return $richTextBox
}
#-------------------------------function--------------------------------------------------------#
function Clean(){
    $MainTopicBox.SelectedItem = "אנא בחר נושא מהרשימה"
    $Description.Clear()
}
function SaveButton{
    if($MainTopicBox.SelectedValue -ne "אנא בחר נושא מהרשימה"){
        if($MainTopicBox.SelectedItem -eq "אחר"){
           $CsvRow = new-object PSObject -property @{
                                   name = $env:USERNAME
                                   Date = (Get-Date).ToString()
                                   Subject = "אחר"
                                   Description = $Description.Text
           }
           $CsvRow|Export-Csv "\\docserver1\SYSTEM_DOCS\CallsStatistics.csv" -Encoding UTF8 -Append
           [void][System.Windows.Forms.MessageBox]::Show("The operation successfully completed :)","Calls Statistics") 
        }
        elseif($SubtopicBox.SelectedValue -eq "אחר"){
               $CsvRow = new-object PSObject -property @{
                                   name = $env:USERNAME
                                   Date = (Get-Date).ToString()
                                   Subject = $MainTopicBox.SelectedItem
                                   Description = $Description.Text
               }
               $CsvRow|Export-Csv "\\docserver1\SYSTEM_DOCS\CallsStatistics.csv" -Encoding UTF8 -Append
               [void][System.Windows.Forms.MessageBox]::Show("The operation successfully completed :)","Calls Statistics") 
            
        }
        else{$CsvRow = new-object PSObject -property @{
                                   name = $env:USERNAME
                                   Date = (Get-Date).ToString()
                                   Subject = $MainTopicBox.SelectedItem
                                   Description = $SubtopicBox.SelectedItem
            }
            $CsvRow|Export-Csv "\\docserver1\SYSTEM_DOCS\CallsStatistics.csv" -Encoding UTF8 -Append
            [void][System.Windows.Forms.MessageBox]::Show("The operation successfully completed :)","Calls Statistics") 
        }
        Clean
    }
    else{[void][System.Windows.Forms.MessageBox]::Show("Please choose a subject from list","Calls Statistics")}
}
#-------------------------------Object--------------------------------------------------------#
$ListSubject=@{}
$Filepath="\\docserver1\SYSTEM_DOCS\files\PowerShell\CallsStatistics.csv"
$CsvFile = Import-Csv -Path $Filepath -Encoding Default
foreach($Param in $CsvFile){
     $ListSubject[$Param.name]=($Param.vaule.Split(",")).trimstart(" ")
}

$MainForm = GenerateForm -title 'Calls Statistics' -Width 270 -Height 220
 $MainForm.minimumSize = New-Object System.Drawing.Size(270,220)
 $MainForm.maximumSize = New-Object System.Drawing.Size(270,220)
 $MainForm.KeyPreview = $True
 $MainForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$FormMain.Close()}})
#-------------------------------topic-------------------------------------#
$MainTopicBox = GenComboBox -data "" -x '140' -y '20' -width '100'
    $MainTopicBoxData = New-Object System.Collections.ArrayList
    $MainTopicBoxData.AddRange([array]"אנא בחר נושא מהרשימה")
    $MainTopicBoxData.AddRange([array]($ListSubject.Keys|Sort-Object))
    $MainTopicBoxData.AddRange([array]"אחר")

#$MainTopicBox.DropDownHeight= 500
$MainTopicBox.IntegralHeight = $True

$MainTopicBox.DataSource =  $MainTopicBoxData
$MainTopicBox.DropDownWidth=($MainTopicBox.DataSource|%{$_.Length}|Sort-Object -Descending|select -First 1) *8
$MainTopicBox.RightToLeft ="yes"
    $MainTopicBox.add_SelectedValueChanged{
        if($MainTopicBox.SelectedValue -eq "אחר"){
                $Description.Enabled = $True
                $SubtopicBox.Enabled=$false
        }
        elseif($MainTopicBox.SelectedValue -ne "אנא בחר נושא מהרשימה"){
                $Description.Enabled = $false
                $SubtopicBox.Enabled=$True
                $Data = New-Object System.Collections.ArrayList
                $Data.AddRange([array]($ListSubject[$MainTopicBox.SelectedValue]|Sort-Object))
                $Data.AddRange([array]"אחר")
                $SubtopicBox.DataSource = $Data
                $SubtopicBox.DropDownWidth=($SubtopicBox.DataSource|%{$_.Length}|Sort-Object -Descending|select -First 1) *7
        }
        else{$Description.Enabled = $false
             $SubtopicBox.Enabled=$false
             $Data = New-Object System.Collections.ArrayList
             $Data.AddRange([array]" ")
             $SubtopicBox.DataSource = $Data
        }
    }

$MainTopicBoxGroup = GenGroupBox -x '130' -y '5' -width '120' -height '45' 
$MainTopicBoxGroup.RightToLeft = "yes"
$MainTopicBoxGroup.Text        = "נושאים"

#-------------------------Subtopic-----------------------------------------#

$SubtopicBox = GenComboBox -data "" -x '15' -y '20' -width '100'
$SubtopicBox.Enabled = $false
$SubtopicBox.RightToLeft ="yes"
$SubtopicBox.add_SelectedValueChanged{
     if($SubtopicBox.SelectedValue -eq "אחר"){
        $Description.Enabled = $True
     }
}

$SubtopicBoxGroup = GenGroupBox -x 5 -y '5' -width '120' -height '45'
$SubtopicBoxGroup.RightToLeft = "yes"
$SubtopicBoxGroup.Text        = "תת-נושאים"

#-------------------------Description-----------------------------------------#
$Description = GeneraterichText -x '10' -y '64' -width '235' -height '50'
$Description.RightToLeft="yes"

$DescriptionGroup = GenGroupBox -x '5' -y '50' -width '245' -height '70' 
$DescriptionGroup.RightToLeft = "yes"

$SaveButton = GenerateButton -x '77' -y '125' -action{SaveButton}
$SaveButton.Image = [System.Convert]::FromBase64String('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
JOXujzP/7Y4z/+2QNf/vmUf/75tK/++bSv/vnEv/75xM/++dTP/vnU3/751N/++dTv/vnU7/755P/++eT//vnk//755P/++dTf/vnEz/75xL/++bSv/vm0r/75lG/+yPNf/miDP/1ns0/7BZKe0+FBRbBAARGP///wCDOg
or13cm5e6QN//tjzX/75tL//Peyv/t4tj/7eHX/+3i1//t4tf/7eLX/+3i1//t4tf/7eLX/+3i1//t4tf/7eLX/+3i1//t4tf/7eLX/+3i1//t4tf/7eHX/+3i2P/z3sr/7ptL/+aJNf/WfTf/sFkr7T4UFFsEABEY////
AII6CivXeCnl7pM8/+2SOv/voFP/9+3j//T19//09PX/9PT1//T09f/09PX/9PT1//T09f/09PX/9PT1//T09f/09PX/9PT1//T09f/09PX/9PT1//T09f/09PX/9PX3//ft4//vn1P/5os6/9Z/PP+wWy3tPhQUWwQAER
j///8AgjkJK9h6LeXulkL/7pVA//CiWP/16+H/7/Dw/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v8PD/9evh/++iWP/njj//14JB/7BcMO0+FBRb
BAARGP///wCCOQgr2Hww5e+aR//umEb/8KVd//Xr4f/v8PD/7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/w8P/16+H/8KRc/+eSRf/XhUb/sF0y7T
4UE1sEABEY////AII5CCvYfjLl751M/++bS//xqGH/9+7k//T09f/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT1//fu5P/wp2H/6JVK/9iHSv+w
XzTtPhQTWwQAERj///8AgjkIK9iANeXwoFL/755Q//GrZv/17OP/7/Dw/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v8PD/9ezj//CqZf/omE//2I
pP/7BgNu0+FBNbBAARGP///wCCOAcr2IE55fCjWP/volb/8a5s//Xs4//v8PD/7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/w8P/17OP/8K1r/+ib
Vf/YjVX/sWE57T4UE1sEABEY////AII4BivZhD7l8ahi//CnYP/ysnT/+O/n//T09f/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT1//jv5//xsX
P/6aBe/9mSXf+xZD3tPhMSWwQAERj///8AgjcFK9qIRuXysG//8a5t//O5gP/27ub/7/Dw/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v8PD/9u7m
//K4f//qp2v/2pho/7FnQ+0+ExJbBAARGP///wCCNgQr2o1O5fO4fv/zt33/9MCN//bv6P/v8PD/7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/w8P
/27+j/9L+N/+uvev/bn3b/smpJ7T4TEVsEABEY//+vAIU3Aircklfk9b+M//S+iv/1x5n/+PLs//P09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/9PT0//T09P/09PT/
8/T0//jy7P/1xpj/7LaH/9ymgv+zblDtQBMRWAMAERb/6C4AjjoBI9iLUN71wpP/9cOU//bLof/28Ov/7+/w/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+
/v7//v7/D/9vDr//bKof/uu5D/3KiI/7FpS+ZFFRBGAAAQDv//RgA9EQAGvGIiceGcZPL1xZb/+M6m//bx7P/w8fL/8PDx//Dw8f/w8PH/8PDx//Dw8f/w8PH/8PDx//Dw8f/w8PH/8PDx//Dw8f/w8PH/8PDx//Dw8f/w
8PH/8PDx//Dx8v/28ez/982m/+27kv/JhmD1ikElhhsCDxoJABEFAAAAAP//CwBMDgAHvGIjcdmNUt/dmmXl3a6M5duvkOXbr5Dl26+Q5duvkOXbr5Dl26+Q5duvkOXbr5Dl26+Q5duvkOXbr5Dl26+Q5duvkOXbr5Dl26
+Q5duvkOXbr5Dl26+Q5d2ujOXbmGTlzoRR4ZpMJH8XAgcZAwARBRUAEwECAAIABQAFAP//AAB7LwAHjzoBI4Y2ASqCMgArgzIAK4MyACuDMgArgzIAK4MyACuDMgArgzIAK4MyACuDMgArgzIAK4MyACuDMgArgzIAK4My
ACuDMgArgzIAK4MyACuDMgArgTEAK4AzAit7LwYnOBMHDgAAEAQUABIBEAAQAAAAAAADAAMAAwADAAIAAgABAAEAAQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgANAAMADQANABABEwASABwAHAARABEAwAAAAYAAAACAAAAAAAAAAIAAAACAAAAAgAAAAIAAAACAAAAAgAAAAIAAAACAAAAAgAAAAIAAAACAAAAAgAAAAIAAAACAAAAAgAAA
AIAAAACAAAAAgAAAAIAAAACAAAAAgAAAAIAAAACAAAAAgAAAAIAAAADAAAAA4AAAAf////c=')
$SaveButton.Size = '40, 40'


$ExitButton=GenerateButton -x '140' -y '125' -action{$MainForm.Close()}
$ExitButton.Image = [System.Convert]::FromBase64String('
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
SUVORK5CYII=')
$ExitButton.Size = '40, 40'
$version = GenerateLabel -text 'V 1.2' -x '2' -y '167'
$version.Font = New-Object System.Drawing.Font("Times New Roman",7,[System.Drawing.FontStyle]::Regular)
$CurrentUser = GenerateLabel -text $env:USERNAME -x '2' -y '157'
$CurrentUser.Font = New-Object System.Drawing.Font("Times New Roman",7,[System.Drawing.FontStyle]::Regular)

#-----------------------------------------------------------------------------#
$FromObject = @($MainTopicBox, $MainTopicBoxGroup,$SubtopicBox,$SubtopicBoxGroup,$Description,$DescriptionGroup,$SaveButton,$ExitButton,$version,$CurrentUser)
foreach($Ctl in $FromObject){$MainForm.Controls.Add($Ctl)}
$MainForm.ShowDialog()