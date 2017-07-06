
#Last modified by zahi ohana 06.07.2017
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$ExSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri '<server>/powershell?serializationLevel=Full'
Import-PSSession $ExSession  
#-------------------------------------function-----------------------------#
function GenerateForm([string]$title, [int]$Width, [int]$Height){
 $form = New-Object System.Windows.Forms.Form
 $form.Text = $title
 $form.Width = $Width
 $form.Height = $Height
 $form.AutoSize = $true
 $form.StartPosition = "CenterScreen"
 $Icon = New-Object system.drawing.icon ("SharedMailbox.ico")
 $form.Icon = $Icon
return $form
}
function Gentabctrl([int]$width, [int]$height,[int]$x, [int]$y,$Selected= '',$tabpage=''){
 $tabctrl = New-Object windows.Forms.TabControl
 $tabctrl.Size = New-Object System.Drawing.Size($width,$height)
 $tabctrl.Location = New-Object System.Drawing.Point($x,$y)
 $tabctrl.SelectedTab = $Selected
 $tabctrl.tabpages.AddRange(@($tabpage))
return $tabctrl
}
function Gentabpage([string]$text,[string]$range){
$tabpage = New-Object windows.Forms.TabPage
$tabpage.Text = $text
return $tabpage
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
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
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
    $textBox.Location = New-Object System.Drawing.Size($x,$y) 
    $textBox.Size = New-Object System.Drawing.Size($width,$height)
    $textBox.Text = $text
    return $textBox
    }
function CheckedListBox([int]$x, [int]$y, [int]$width, [int]$height){
    $listBox1 = New-Object System.Windows.Forms.CheckedListBox
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = $width
        $System_Drawing_Size.Height = $height
    $listBox1.Size = $System_Drawing_Size
    $listBox1.DataBindings.DefaultDataSourceUpdateMode = 0
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = $x
        $System_Drawing_Point.Y = $y
    $listBox1.Location = $System_Drawing_Point
    return $listBox1
}
function GenerateListBox([int]$x, [int]$y, [int]$width, [int]$height){
    $listBox1 = New-Object System.Windows.Forms.ListBox
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = $width
        $System_Drawing_Size.Height = $height
    $listBox1.Size = $System_Drawing_Size
    $listBox1.DataBindings.DefaultDataSourceUpdateMode = 0
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = $x
        $System_Drawing_Point.Y = $y
    $listBox1.Location = $System_Drawing_Point
    return $listBox1
}
function GenRadioBox([int]$x, [int]$y, [int]$width, [int]$height,[string]$text){
 $radiobutton = New-Object System.Windows.Forms.RadioButton
 $radiobutton.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
 $radiobutton.Location = New-Object System.Drawing.Size($x,$y)
 $radiobutton.Size = New-Object System.Drawing.Size($width,$height)
 $radiobutton.TabStop = $True
 $radiobutton.Text = $text
 return $radiobutton
 }
function GenComboBox([array]$data, [int]$x, [int]$y, [int]$width, [int]$height){
    $ComboBox = New-Object System.Windows.Forms.ComboBox
    $ComboBox.DataSource = @($data)
    $ComboBox.Location  = New-Object System.Drawing.Point($x,$y)
    $ComboBox.Size = New-Object System.Drawing.Size($width,$height)
    return $ComboBox
}
$ErrorProvider = New-Object System.Windows.Forms.ErrorProvider
#-----------------------function tab new--------------------------------#
function CreateMail(){
$errorProvider.Clear()
if(!($NewNameINPUT.Text -eq "")){
      CheckAvailable
      if(!($NewNameINPUT.Text -eq "Sorry ,the name isn't available")){
       if(!($NewAliasINPUT.Text -eq "")){
            CheckAvailable2
            if(!($NewAliasINPUT.Text -eq "Sorry ,the name isn't available")){
                if(!($newWhomanageINPUT.text -eq "")){
                $mailname = $NewAliasINPUT.Text + "@<Domain>"
                $desname ="מנוהל על ידי"+" "+$newWhomanageINPUT.text
                $newmailbox =New-Mailbox -Name $NewNameINPUT.Text -SamAccountName $NewAliasINPUT.Text -Alias $NewAliasINPUT.Text `
                             -OrganizationalUnit "<ou>" -Database $NewDbcombobox.SelectedItem `
                             -UserPrincipalName $mailname -Shared 
                Start-Sleep -Seconds 7
                
                $Question = [Microsoft.VisualBasic.Interaction]::MsgBox("The mailbox was created, do you want to add members?",36, "Success")
                if($Question -eq "yes"){
                    Set-ADUser $NewAliasINPUT.Text -Description $desname
                    Set-MailboxSentItemsConfiguration $NewAliasINPUT.Text -SendAsItemsCopiedTo From -SendOnBehalfOfItemsCopiedTo From
                    ClearForm
                    $Tabctrl.SelectTab($AddTAB)
                    $AddAliasINPUT.Text = $NewAliasINPUT.Text
                }
                if($Question -eq "no"){
                    Set-ADUser $NewAliasINPUT.Text -Description $desname
                    Set-MailboxSentItemsConfiguration $NewAliasINPUT.Text -SendAsItemsCopiedTo From -SendOnBehalfOfItemsCopiedTo From
                    }
                    clearNewMail
                }
                else{$errorprovider.BlinkStyle ="NeverBlink"
                     $errorprovider.SetIconPadding($newWhomanageINPUT, 5)
                     $ErrorProvider.SetError($newWhomanageINPUT, "The description is missing, please write description.")}
            }
            else{return}
       } 
       else{$errorprovider.BlinkStyle ="NeverBlink"
            $errorprovider.SetIconPadding($NewAliasINPUT, 2.5)
            $ErrorProvider.SetError($NewAliasINPUT, "The alias is missing, please write alias.")}
    }
      else{return}
}
else{$errorprovider.BlinkStyle ="NeverBlink"
     $errorprovider.SetIconPadding($NewNameINPUT, 2.5)
     $errorprovider.SetError($NewNameINPUT, "The name missing, please write name")}
}
function clearNewMail(){
 $errorProvider.Clear()
 $NewNameINPUT.text = ""
 $NewNameINPUT.BackColor = "window"
 $NewAliasINPUT.text = ""
 $NewAliasINPUT.BackColor = "window"
 $newWhomanageINPUT.text = ""
 $NewDbcombobox.Text =$newDbData.name[0]
}
function CheckAvailable(){
if($NewNameINPUT.Text -ne ""){
    $mailchack = Get-Mailbox $NewNameINPUT.Text
    if(!($mailchack)){$NewNameINPUT.BackColor ="#22FD00" <#green#>
    }
    else{$NewNameINPUT.BackColor ="#FF0000" <#red#>
         $NewNameINPUT.TextAlign = "center"
         $NewNameINPUT.Text = "Sorry ,the name isn't available"}
    }
}
function CheckAvailable2(){
if($NewNameINPUT.Text -ne ""){
    $mailchack = Get-Mailbox $NewAliasINPUT.Text
    if(!($mailchack)){$NewAliasINPUT.BackColor ="#22FD00" <#green#>
    }
    else{$NewAliasINPUT.BackColor ="#FF0000" <#red#>
         $NewAliasINPUT.TextAlign = "center"
         $NewAliasINPUT.Text = "Sorry ,the name isn't available"}
    }
}
#---------------------------------------function tab add-------------------------------------#
function SearchAD2($Alias){
    $result = @()
    $result += Get-ADUser -Filter {Name -like $Alias}|Select-Object name,samaccountname
    $resultsam = Get-ADUser -Filter { SamAccountName -eq $Alias }
    $resultdispaly=Get-ADUser -Properties DisplayName -Filter{DisplayName -like $Alias}|Select-Object name,samaccountname
    if($result.samaccountname -gt "1"){
       $DisplayResult=@()
       foreach ($Items in $Result){
       $DisplayResult+=$Items.SamAccountName+" - "+$Items.name   
    }
        return $DisplayResult}
    if($resultdispaly.samaccountname -gt "1"){
        foreach ($Items in $Result){
          $DisplayResult+=$Items.SamAccountName+" - "+$Items.name }
     }
}
function SearchAD($Alias){
    $result = @()
    $result += Get-ADUser -Filter {Name -like $Alias}|Select-Object name,samaccountname
    $resultsam = Get-ADUser -Filter { SamAccountName -eq $Alias }
    $resultdispaly=Get-ADUser -Properties DisplayName -Filter{DisplayName -like $Alias}|Select-Object name,samaccountname
    if($result.samaccountname -gt "1"){
     return $result.SamAccountName}
     if($resultdispaly.samaccountname -gt "1"){
     return $resultdispaly.SamAccountName}
}
function addshowcontrol(){
foreach($addshowcontrols in $AddTABshowControl){
     $AddTAB.Controls.Add($addshowcontrols)}
}
function addshowcontrolCalendar(){
foreach($addshowcontrols in $AddcalanderControl){
     $AddTAB.Controls.Add($addshowcontrols)}
}
function FullAcess(){
 if(($AddAliasINPUT.Text -notlike "Mail not found for*")-and($AddAliasINPUT.text -ne "Alias Name Only")-and($AddAliasINPUT.text -ne "Plase fill an alias name")`
 -and($AddAliasINPUT.Text -ne "Fax Number Not Found")){
  ClearForm2
  if($AddAliasINPUT.Text -match $hebrewname -or $AddAliasINPUT.Text -match $AllPat){
       $hebrewmail = $AddAliasINPUT.Text + "*"
       $SearchHebrewmail = SearchAD $hebrewmail
       $ResultPLus=@()
       foreach ($item in $SearchHebrewmail){
           $result=Get-ADUser $item |Select-Object name,SamAccountName
           $ResultPLus+=$result.name+" - "+$result.SamAccountName
       }
       if($SearchHebrewmail.Count -gt "1"){
              aliasfrom
              if (!($aliasLST.SelectedItem.count -eq "0")){
                    $Split=$aliasLST.SelectedItem.Split("-")
                    $Splituser=$Split[1].TrimStart(" ")
                    $AddAliasINPUT.Text=$Splituser}
                     else{return}           }
       elseif($SearchHebrewmail.Count -eq "1"){
          $Addlist= Get-Mailbox $SearchHebrewmail|Select-Object PrimarySmtpAddress
          $AddAliasINPUT.Text=$Addlist.PrimarySmtpAddress}
       else{$Formherew.hide()
            $AddAliasINPUT.TextAlign = "center"
            $AddAliasINPUT.BackColor ="red"
            $user=$AddAliasINPUT.text
            $AddAliasINPUT.Text = "Mail not found for $user"
            $AddFullRadio.Checked = $false
            return
       }
}
  If($AddAliasINPUT.text -match $EnglishName -or $AddAliasINPUT.Text -match $AllPat){
    $Checkuser=(Get-Mailbox $AddAliasINPUT.text |Select-Object WindowsEmailAddress).WindowsEmailAddress
 if($Checkuser -ne $null){
      $AddAliasINPUT.textalign="center"
      $AddAliasINPUT.text=$Checkuser
      $AddAliasINPUT.Enabled = $false
      $AddMembersFull = Get-Mailbox $AddAliasINPUT.Text |get-MailboxPermission |Select-Object user,AccessRights | where { ($_.AccessRights -contains "FullAccess")}
      $addMembersName=$AddMembersFull.user|where{(<Users Default>)}
     
      $AddMembers = $addMembersName.split("\") |  where{($_ -ne <Domain>)} 
      $AddOnlyMembers= @()
        foreach($AddMembersfu in $AddMembers){
            try{$AddOnlyMembers += get-ADUser $AddMembersfu -Properties * |Select-Object Name,uid| where{$_.uid -ne $null}}
            catch{$error1= "all user"+$AddOnlyMembers}}    
     if($AddOnlyMembers.Count -gt "0"){
        addshowcontrol
          $addNamelist = $AddOnlyMembers.name |Sort-Object 
           foreach ($item2 in $addNamelist){
             $AddListMemebers.Items.Add($item2)}
     }
     else{$AddMsgBox =[Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this mailbox, Do you wish to continue?",36, "Error Permission")
        if($AddMsgBox -eq "yes"){addshowcontrol 
           $AddAliasINPUT.Enabled = $false}
        else{ClearForm}
     }
    }
 else{$AddAliasINPUT.TextAlign = "center"
      $AddAliasINPUT.BackColor ="red"
      $User=$AddAliasINPUT.Text
      $AddAliasINPUT.Text = "Mail not found for $user"
      $AddFullRadio.Checked = $false
 }
 return
}
  if($AddAliasINPUT.Text -match $pat){
   $faxother = "*"+$AddAliasINPUT.Text +"*"
   $faxuser =Get-ADUser -filter {otherFacsimileTelephoneNumber -like $faxother}|Select-Object SamAccountName
    if($faxuser){
       if($faxuser.SamAccountName.count -eq "1"){
        $AddAliasINPUT.Text = $faxuser.SamAccountName}
       else{[Microsoft.VisualBasic.Interaction]::MsgBox("More than one user were found.
Please write down the user's Alias name.",16, "Error")
           }
               }
    else{$AddAliasINPUT.TextAlign = "center"
         $AddAliasINPUT.BackColor ="red"
         $AddAliasINPUT.Text = "Fax number not found"
         $AddFullRadio.Checked = $false
         return
    }
                                 }
 else{$AddAliasINPUT.TextAlign = "center"
      $AddAliasINPUT.BackColor ="red"
      $User=$AddAliasINPUT.Text
      $AddAliasINPUT.Text = "Mail not found for $user"
      $AddFullRadio.Checked = $false}
}
 else{$AddAliasINPUT.TextAlign = "center"
     $AddAliasINPUT.BackColor ="red"
     $AddAliasINPUT.Text = "Plase fill an alias name"
     $AddFullRadio.Checked = $false}
}  
function SendAS(){
 if(($AddAliasINPUT.Text -notlike "Mail not found for*")-and($AddAliasINPUT.text -ne "Alias Name Only")-and($AddAliasINPUT.text -ne "Plase fill an alias name")`
 -and($AddAliasINPUT.Text -ne "Fax Number Not Found")){
  ClearForm2
  if($AddAliasINPUT.Text -match $pat){
   $faxother = "*"+$AddAliasINPUT.Text +"*"
   $faxuser =Get-ADUser -filter {otherFacsimileTelephoneNumber -like $faxother}|Select-Object SamAccountName
    if($faxuser){
       if($faxuser.SamAccountName.count -eq "1"){
        $AddAliasINPUT.Text = $faxuser.SamAccountName}
       else{[Microsoft.VisualBasic.Interaction]::MsgBox("More than one user were found.
Please write down the user's Alias name.",16, "Error")
           }
               }
    else{$AddAliasINPUT.TextAlign = "center"
         $AddAliasINPUT.BackColor ="red"
         $AddAliasINPUT.Text = "Fax number not found"
         $AddSendASRadio.Checked = $false
         return
    }
                                   }
  if($AddAliasINPUT.Text -match $hebrewname){
       $hebrewmail = $AddAliasINPUT.Text + "*"
       $SearchHebrewmail = SearchAD $hebrewmail
       $ResultPLus=@()
       foreach ($item in $SearchHebrewmail){
           $result=Get-ADUser $item |Select-Object name,SamAccountName
           $ResultPLus+=$result.name+" - "+$result.SamAccountName
       }
       if($SearchHebrewmail.Count -gt "1"){
              aliasfrom
              if (!($aliasLST.SelectedItem.count -eq "0")){
                    $Split=$aliasLST.SelectedItem.Split("-")
                    $Splituser=$Split[1].TrimStart(" ")
                    $AddAliasINPUT.Text=$Splituser}
                     else{return}           }
       elseif($SearchHebrewmail.Count -eq "1"){
              $Addlist= Get-Mailbox $SearchHebrewmail|Select-Object PrimarySmtpAddress
              $AddAliasINPUT.Text=$Addlist.PrimarySmtpAddress}
       else{$Formherew.hide()
            $AddAliasINPUT.TextAlign = "center"
            $AddAliasINPUT.BackColor ="red"
            $user=$AddAliasINPUT.Text
            $AddAliasINPUT.Text = "Mail not found for $user"
            $AddSendASRadio.Checked = $false
            return
       }
}
  if($AddAliasINPUT.text -match $EnglishName){
   $UserChack=(Get-Mailbox $AddAliasINPUT.Text|Select-Object WindowsEmailAddress).WindowsEmailAddress
 if($UserChack -ne $null){
      $AddAliasINPUT.textalign="center"
      $AddAliasINPUT.text=$UserChack
      $AddAliasINPUT.Enabled = $false
    $AddTAB.Controls.Add($Addlodiangsendas)
    $AddTAB.Cursor ="WaitCursor"
 if($addSenASMembers= Get-Mailbox $AddAliasINPUT.Text | Get-ADPermission | where {($_.ExtendedRights -like “*Send-As*”) -and -not ($_.User -like “NT AUTHORITY\SELF”)}|Select-Object user){
    $addSenASMember = $addSenASMembers.user
    $addSenASSplit = $addSenASMember.split("\") |  where{$_ -ne <Domain>} 
    $addMemerSendasheb=@()
    $AddTAB.Controls.remove($Addlodiangsendas)
      addshowcontrol
      foreach ($AddMembersSendas in $addSenASSplit){
       $addMemerSendasheb += Get-ADUser $AddMembersSendas|Select-Object name|Sort-Object}
       foreach ($AddMembersSendasjoin in $addMemerSendasheb.name){
                $AddListMemebers.Items.Add($AddMembersSendasjoin)}
                  $AddTAB.Cursor ="Arrow"}
 else{$AddTAB.Controls.remove($Addlodiangsendas)
       $AddTAB.Cursor ="Arrow"
       $AddMsgBox =[Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this mailbox, Do you wish to continue?",36, "Error Permission")
        if($AddMsgBox -eq "yes"){$AddAliasINPUT.Enabled = $false
        addshowcontrol}
         else{ClearForm}
        }
 }
 else{$AddAliasINPUT.TextAlign = "center"
     $AddAliasINPUT.BackColor ="red"
     $user=$AddAliasINPUT.Text
     $AddAliasINPUT.Text = "Mail not found for $User"
     $AddSendASRadio.Checked = $false
  }
}
  else{$AddAliasINPUT.TextAlign = "center"
     $AddAliasINPUT.BackColor ="red"
     $user=$AddAliasINPUT.Text
     $AddAliasINPUT.Text = "Mail not found for $User"
     $AddSendASRadio.Checked = $false}
 } 
 else{$AddAliasINPUT.TextAlign = "center"
      $AddAliasINPUT.BackColor ="red"
      $AddAliasINPUT.Text = "Plase fill an alias name"
      $AddSendASRadio.Checked = $false}
}
function SendOn(){ 
 if(($AddAliasINPUT.Text -notlike "Mail not found for*")-and($AddAliasINPUT.text -ne "Alias Name Only")-and($AddAliasINPUT.text -ne "Plase fill an alias name")`
 -and($AddAliasINPUT.Text -ne "Fax Number Not Found")){
  ClearForm2
  if($AddAliasINPUT.Text -match $pat){
   $faxother = "*"+$AddAliasINPUT.Text +"*"
   $faxuser =Get-ADUser -filter {otherFacsimileTelephoneNumber -like $faxother}|Select-Object SamAccountName
    if($faxuser){
       if($faxuser.SamAccountName.count -eq "1"){
        $AddAliasINPUT.Text = $faxuser.SamAccountName}
       else{[Microsoft.VisualBasic.Interaction]::MsgBox("More than one user were found.
Please write down the user's Alias name.",16, "Error")
           }
               }
    else{$AddAliasINPUT.TextAlign = "center"
         $AddAliasINPUT.BackColor ="red"
         $AddAliasINPUT.Text = "Fax number not found"
         $AddSendonRadio.Checked = $false
         return
    }
                     }
  if($AddAliasINPUT.Text -match $hebrewname){
       $hebrewmail = $AddAliasINPUT.Text + "*"
       $SearchHebrewmail = SearchAD $hebrewmail
       $ResultPLus=@()
       foreach ($item in $SearchHebrewmail){
           $result=Get-ADUser $item |Select-Object name,SamAccountName
           $ResultPLus+=$result.name+" - "+$result.SamAccountName
       }
       if($SearchHebrewmail.Count -gt "1"){
              aliasfrom
              if (!($aliasLST.SelectedItem.count -eq "0")){
                    $Split=$aliasLST.SelectedItem.Split("-")
                    $Splituser=$Split[1].TrimStart(" ")
                    $AddAliasINPUT.Text=$Splituser}
                     else{return}           }
       elseif($SearchHebrewmail.Count -eq "1"){
              $Addlist= Get-Mailbox $SearchHebrewmail|Select-Object PrimarySmtpAddress
              $AddAliasINPUT.Text=$Addlist.PrimarySmtpAddress}
       else{$Formherew.hide()
            $AddAliasINPUT.TextAlign = "center"
            $AddAliasINPUT.BackColor ="red"
            $user=$AddAliasINPUT.Text
            $AddAliasINPUT.Text = "Mail not found for $user"
            $AddSendonRadio.Checked = $false
            return
       }
}
  if($AddAliasINPUT.text -match $EnglishName){
     $AddsendOnMem = Get-Mailbox $AddAliasINPUT.Text |Select-Object GrantSendOnBehalfTo, WindowsEmailAddress
     if($AddsendOnMem -ne $null){
     $AddAliasINPUT.textalign="center"
     $AddAliasINPUT.Text=$AddsendOnMem.WindowsEmailAddress
        $AddAliasINPUT.Enabled = $false
     if($AddsendOnMem.GrantSendOnBehalfTo -gt "0"){
       $AddsendOnMemnot= $AddsendOnMem.GrantSendOnBehalfTo| Where {$_ -notlike <domain>} 
       $AddSendoName =@()
         foreach($item in $AddsendOnMemnot){
                  $AddSendoName += Get-Mailbox $item|Select-Object name}
                    $AddSendoNamesoft =$AddSendoName.name|Sort-Object
                    addshowcontrol
         foreach ($item2 in $AddSendoNamesoft){
                  $AddListMemebers.Items.Add($item2)} 
                  $AddTAB.Cursor ="Arrow"}
    else{$AddMsgBox = [Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this mailbox, Do you wish to continue?",36, "Error Permission")
        if($AddMsgBox -eq "yes"){addshowcontrol
        $AddAliasINPUT.Enabled = $false
                              }
        else{ClearForm}
        }
                    }
      else{$AddAliasINPUT.TextAlign = "center"
           $AddAliasINPUT.BackColor ="red"
           $user=$AddAliasINPUT.Text
           $AddAliasINPUT.Text = "Mail not found for $User"
           $AddSendOnRadio.Checked = $false
     }
  }
  else{$AddAliasINPUT.TextAlign = "center"
       $AddAliasINPUT.BackColor ="red"
       $user=$AddAliasINPUT.Text
       $AddAliasINPUT.Text = "Mail not found for $User"
       $AddSendOnRadio.Checked = $false
  }
 }
 else{$AddAliasINPUT.TextAlign = "center"
      $AddAliasINPUT.BackColor ="red"
      $AddAliasINPUT.Text = "Plase fill an alias name"
      $AddSendonRadio.Checked = $false}
}
function Calendarfull(){
 if(($AddAliasINPUT.Text -notlike "Mail not found for*")-and($AddAliasINPUT.text -ne "Alias Name Only")-and($AddAliasINPUT.text -ne "Plase fill an alias name")){
   $AddListMemebersCalendar.Items.Clear()
   $AddListMemebersCalendarImport.Items.Clear()
   ClearForm2
   if($AddAliasINPUT.Text -match $pat){
      $faxother = "*"+$AddAliasINPUT.Text +"*" 
      $faxuser =Get-ADUser -filter {otherFacsimileTelephoneNumber -like $faxother}|Select-Object SamAccountName
      if($faxuser){
       if($faxuser.SamAccountName.count -eq "1"){
        $AddAliasINPUT.Text = $faxuser.SamAccountName
       }
       else{[Microsoft.VisualBasic.Interaction]::MsgBox("More than one user were found.
Please write down the user's Alias name.",16, "Error")
       }
               }
      else{$AddAliasINPUT.TextAlign = "center"
           $AddAliasINPUT.BackColor ="red"
           $AddAliasINPUT.Text = "Fax number not found"
           $AddSendASRadio.Checked = $false
           return
      }
   }
   if($AddAliasINPUT.Text -match $hebrewname){
       $hebrewmail = $AddAliasINPUT.Text + "*"
       $SearchHebrewmail = SearchAD $hebrewmail
       $ResultPLus=@()
       foreach ($item in $SearchHebrewmail){
           $result=Get-ADUser $item |Select-Object name,SamAccountName
           $ResultPLus+=$result.name+" - "+$result.SamAccountName
       }
       if($SearchHebrewmail.Count -gt "1"){
          aliasfrom
          if(!($aliasLST.SelectedItem.count -eq "0")){
               $Split=$aliasLST.SelectedItem.Split("-")
               $Splituser=$Split[1].TrimStart(" ")
               $AddAliasINPUT.Text=$Splituser
          }
          else{return}           
       }
       elseif($SearchHebrewmail.Count -eq "1"){
              $Addlist= Get-Mailbox $SearchHebrewmail|Select-Object PrimarySmtpAddress
              $AddAliasINPUT.Text=$Addlist.PrimarySmtpAddress
       }
       else{$Formherew.hide()
            $AddAliasINPUT.TextAlign = "center"
            $AddAliasINPUT.BackColor ="red"
            $AddAliasINPUT.Text = "Mail not found for $SearchHebrewmail"
            $AddcalendarRadio.Checked = $false}
}
   If($AddAliasINPUT.text -match $EnglishName){
   $mailboxname = get-mailbox $AddAliasINPUT.text |Select-Object SamAccountName, WindowsEmailAddress
   #MailDisplay
    if($mailboxname -ne $NUll){
       $AddAliasINPUT.textalign="center"
       $AddAliasINPUT.text=$mailboxname.WindowsEmailAddress
       $AddAliasINPUT.Enabled = $false
       $usernamemail= $mailboxname.SamAccountName
       $calendarheb = get-MailboxFolderPermission ("{0}:\לוח שנה" -f $mailboxname.SamAccountName)
       if($calendarheb){
          $calendarheb1= get-MailboxFolderPermission ("{0}:\לוח שנה" -f $mailboxname.SamAccountName)|Select-Object user,AccessRights|where {($_.user -ne "Default") -and ($_.user -ne "Anonymous")}
          if($calendarheb1 -ne $NUll){
             addshowcontrolCalendar
             foreach ($addcalendar in $calendarheb1){
                      $showpermi= $addcalendar.user+" - "+$addcalendar.AccessRights
                      $AddListMemebersCalendar.Items.Add($showpermi) 
             }
          }
          else{$AddMsgBox =[Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this calendar, Do you wish to continue?",36, "Error Permission")
               if($AddMsgBox -eq "yes"){
                  addshowcontrolCalendar 
                  $AddAliasINPUT.Enabled = $false
               }
               else{ClearForm}
          }
                                       
        }
       elseif($calendarheb2=get-MailboxFolderPermission ("{0}:\calendar" -f $mailboxname.SamAccountName)){
              $calendareng = get-MailboxFolderPermission  ("{0}:\calendar" -f $mailboxname.SamAccountName)|Select-Object user,AccessRights |where {($_.user -ne "Default") -and ($_.user -ne "Anonymous")}
              if($calendareng.count -ne "0"){
                 addshowcontrolCalendar
                 foreach ($calendardisplay in $calendareng){
                          $showpermi= $calendardisplay.user+" - "+$calendardisplay.AccessRights
                          $AddListMemebersCalendar.Items.Add($showpermi)
                 }
              }
              else{$AddMsgBox =[Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this calendar, Do you wish to continue?",36, "Error Permission")
                   if($AddMsgBox -eq "yes"){
                      addshowcontrolCalendar 
                      $AddAliasINPUT.Enabled = $false
                   }
                   else{ClearForm}
              }
       }
       else{$AddAliasINPUT.TextAlign = "center"
            $AddAliasINPUT.BackColor ="red"
            $AddAliasINPUT.Text = "Mail not found"
            $AddSendOnRadio.Checked = $false
       }
       }
    else{$AddAliasINPUT.TextAlign = "center"
         $AddAliasINPUT.BackColor ="red"
         $user=$AddAliasINPUT.Text
         $AddAliasINPUT.Text = "Mail not found for $user"
         $AddcalendarRadio.Checked = $false
    }
}
   else{$AddAliasINPUT.TextAlign = "center"
        $AddAliasINPUT.BackColor ="red"
        $user=$AddAliasINPUT.Text
        $AddAliasINPUT.Text = "Mail not found for $user"
        $AddcalendarRadio.Checked = $false
   }
  }
 else{$AddAliasINPUT.TextAlign = "center"
      $AddAliasINPUT.BackColor ="red"
      $AddAliasINPUT.Text = "Plase fill an alias name"
      $AddcalendarRadio.Checked = $false
  }
}
function ClearForm(){
$AddAliasINPUT.BackColor = "window"
 $AddAliasINPUT.Text = "Alias Name Only"
 $AddAliasINPUT.TextAlign ="center"
 $AddAliasINPUT.Enabled = $true
 $AddFullRadio.Checked= $false
 $AddSendASRadio.Checked= $false
 $AddSendOnRadio.Checked= $false
 $AddcalendarRadio.Checked= $false
 $AddListMemebersImport.Items.Clear()  
 $AddUserNameINPUT.Text = ""
 $AddListMemebers.Items.Clear() 
 $AddListRemove.Items.Clear()
 $AddListMemebersCalendarImport.Items.Clear()
 $AddListRemoveCalendar.Items.Clear()
 $AddTAB.Controls.remove($AddListMemebersImport)
 foreach($tabremove in $AddTABshowControl){
     $addTAB.Controls.Remove($tabremove)}
 foreach($tabremove in $AddcalanderControl){
     $addTAB.Controls.Remove($tabremove)}   
}
function ClearForm2(){
 $AddAliasINPUT.BackColor = "window"
 $AddUserNameINPUT.BackColor = "window"
 $AddListMemebersImport.Items.Clear()  
 $AddUserNameINPUT.Text = ""
 $AddListMemebers.Items.Clear() 
 $AddListRemove.Items.Clear()
 $AddTAB.Controls.remove($AddListMemebersImport)
 foreach($tabremove in $AddTABshowControl){
     $addTAB.Controls.Remove($tabremove)}
 foreach($tabremove in $AddcalanderControl){
     $addTAB.Controls.Remove($tabremove)}   
}
function ClearForm3(){
 $AddAliasINPUT.Enabled = $true
 $AddFullRadio.Checked= $false
 $AddSendASRadio.Checked= $false
 $AddSendOnRadio.Checked= $false
 $AddcalendarRadio.Checked= $false
 $AddListMemebersImport.Items.Clear()  
 $AddAliasINPUT.Enabled = $True
 $AddListMemebers.Items.Clear() 
 $AddListRemove.Items.Clear()
 $AddListRemoveCalendar.Items.Clear()
 $AddListMemebersCalendarImport.Items.Clear()
 $AddListMemebersCalendarImport.Items.Clear()
 $AddTAB.Controls.remove($AddListMemebersImport)
 foreach($tabremove in $AddTABshowControl){
     $addTAB.Controls.Remove($tabremove)}
 foreach($tabremove in $AddcalanderControl){
 $addTAB.Controls.Remove($tabremove)}   
}  
function fromhebrew(){
if($MsgUserNameINPUT.Text -ne ""){
   $heblist.Items.Clear()
   foreach($addhebrewname in $ResultPLus){
           $heblist.Items.Add($addhebrewname)
   }
   $Formherew.ShowDialog()
   $MsgUserNameINPUT.Text = ""
}
elseif($AddUserNameINPUT.Text -ne ""){
        $heblist.Items.Clear()
        foreach($addhebrewname in $ResultPLus){
                 $heblist.Items.Add($addhebrewname)}
        $Formherew.ShowDialog()
    }
else{return}
}
function aliasfrom(){
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
   if($MsgAliasINPUT.Text -ne "Alias Name Only"){
      $AliasLST.Items.Clear()
      foreach($addhebrewname in  $ResultPLus){
              $aliasLST.Items.Add($addhebrewname)
      }
      $Formaliashebrew.ShowDialog()    
    }
    else{return}
 }
 elseif($AddAliasINPUT.Text -ne "Alias Name Only"){
   $aliasLST.Items.Clear()
   foreach($addhebrewname in  $ResultPLus){
           $aliasLST.Items.Add($addhebrewname)
   }
   $Formaliashebrew.ShowDialog()
 }
 else{return}
}
function ChooseName(){
if($heblist.SelectedItem.count -eq "1" ){
   $Formherew.hide();$AddUserNameINPUT.Text = ""}
elseif($aliasLST.SelectedItem.count -gt "0"){$Formaliashebrew.hide()}
else{[Microsoft.VisualBasic.Interaction]::MsgBox("No username was selected, Please select",16, "Error")}
}
function addUsrMail(){
$errorProvider.Clear()
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
    #if match hebrew name
    if($MsgUserNameINPUT.Text -match $hebrewname){
       $adduserfilter = $MsgUserNameINPUT.Text + "*"
       $SearchHebName = SearchHeb $adduserfilter| Sort-Object -Unique
       if($SearchHebName.Count -gt "1"){
          $ResultPLus=@()
          foreach ($item in $SearchHebName){
               try{$Result=Get-ADUser $item |Select-Object name,SamAccountName
                   $ResultPLus+=$result.name #+" - "+$result.SamAccountName
               }
               catch{$Result=Get-DistributionGroup "$item"|Select-Object name,SamAccountName
                     $ResultPLus+=$result.name #+" - "+$result.SamAccountName
               }
          }
          fromhebrew
          $MsgUserNameINPUT.Text=""
          $MsgListMemebersImport.Items.Add($heblist.SelectedItem)
          return
       }
       elseif($SearchHebName.Count -eq "1"){
             $UserMailbox = (Get-Mailbox $SearchHebName|Select-Object name).name
             if($UserMailbox -ne $null){
              $MsgListMemebersImport.Items.Add($UserMailbox)
              $MsgUserNameINPUT.Text=""
             }
             elseif($DistributionMailbox = (Get-DistributionGroup "$SearchHebName"|Select-Object name).name){
                    $MsgListMemebersImport.Items.Add($DistributionMailbox)
                    $MsgUserNameINPUT.Text=""   
             }
             else{$Formherew.hide()
                  $MsgUserNameINPUT.BackColor ="Red"
                  $MsgUserNameINPUT.TextAlign = "Center"
                 $MsgUserNameINPUT.Text = "User not found"
            }
       }
       else{$Formherew.hide()
            $MsgUserNameINPUT.BackColor ="Red"
            $MsgUserNameINPUT.TextAlign = "Center"
            $MsgUserNameINPUT.Text = "User not found"
       }
      }
    #if not match hebrewname
    elseif($MsgUserNameINPUT.Text -match $EnglishName){
     $ChceckUser=(Get-Mailbox $MsgUserNameINPUT.Text |Select-Object name,SamAccountName).name
     if($ChceckUser -ne $null){
        $MsgListMemebersImport.Items.Add($ChceckUser)
        $MsgUserNameINPUT.Text= ""
        return
     }
     elseif($DistributionMailbox = (Get-DistributionGroup $MsgUserNameINPUT.Text|Select-Object name).name){
            $MsgListMemebersImport.Items.Add($DistributionMailbox)
            $MsgUserNameINPUT.Text=""   
     }
     else{$Formherew.hide()
          $MsgUserNameINPUT.BackColor ="Red"
          $MsgUserNameINPUT.TextAlign = "Center"
          $MsgUserNameINPUT.Text = "User not found"
     }
    }
    #if Blnk
    elseif($MsgUserNameINPUT.Text -eq ""){
           $Errorprovider.BlinkStyle ="NeverBlink"
           $Errorprovider.SetIconPadding($MsgUserNameINPUT, -15)
           $ErrorProvider.SetError($MsgUserNameINPUT, "write User Name.")
        }
    else{$Formherew.hide()
         $MsgUserNameINPUT.BackColor ="red"
         $MsgUserNameINPUT.TextAlign = "center"
         $MsgUserNameINPUT.text = "User not found"
    }
     
 }
 elseif($AddcalendarRadio.Checked -eq $true){
        $choosepermissionslist.SelectedItems.Clear()
        #if match hebrew name
        if($AddUserNameINPUT.Text -match $hebrewname){
            $adduserfilter = $AddUserNameINPUT.Text + "*"
            $SearchHebName = searchAD $adduserfilter
            $ResultPLus=@()
            foreach ($item in $SearchHebName){
                     $result=Get-ADUser $item |Select-Object name,SamAccountName
                     $ResultPLus+=$result.name+" - "+$result.SamAccountName
            }
            if($SearchHebName.Count -gt "1"){
               fromhebrew
               if(!($heblist.SelectedItem.count -eq "0")){
                    $CalendarMain.ShowDialog()
                    if(!($choosepermissionslist.SelectedItem.count -eq "0")){
                         $Split=$heblist.SelectedItem.Split(" - ")
                         $Splituser=$Split[4].TrimStart(" ")
                         $Calendarhewbrewname =Get-Mailbox $Splituser |Select-Object name
                         if($Calendarhewbrewname -eq $null){[Microsoft.VisualBasic.Interaction]::MsgBox("The mailbox "+ $Splituser + " not found",0, "Error")
                         return}
                         $Calendarname = $Calendarhewbrewname.name+" <- "+$choosepermissionslist.SelectedItem
                         $AddListMemebersCalendarImport.Items.Add($Calendarname)
                         return
                 } 
                    else{return}    
                 }
               else{return}
        }
        elseif($SearchHebName.Count -eq "1"){
                $CalendarMain.ShowDialog()
            if(!($choosepermissionslist.SelectedItem.count -eq "0")){
                 $Calendarhe = Get-Mailbox $SearchHebName|Select-Object name
                 $Calendarname = $Calendarhe.name+" <- "+$choosepermissionslist.SelectedItem
                 $AddListMemebersCalendarImport.Items.Add($Calendarname)
                 return 
            } 
            else{return}
        }
        else{$Formherew.hide()
             $CalendarMain.Close()
             $AddUserNameINPUT.BackColor ="red"
             $AddUserNameINPUT.TextAlign = "center"
             $AddUserNameINPUT.Text = "User not found"}
      }
        #if not match hebrewname
        elseif($addname =Get-Mailbox $AddUserNameINPUT.Text |Select-Object name,SamAccountName){
               $CalendarMain.ShowDialog()
            if(!($choosepermissionslist.SelectedItem.count -eq "0")){
                 $Calendarhewbrewname =get-ADUser $AddUserNameINPUT.Text |Select-Object name
                 if($Calendarhewbrewname){$Calendarname = $Calendarhewbrewname.name+" <- "+$choosepermissionslist.SelectedItem
                                          $AddListMemebersCalendarImport.Items.Add($Calendarname)
                                          $AddUserNameINPUT.Text= ""
                                          return
                 }
                 else{$AddUserNameINPUT.BackColor ="red"
                      $AddUserNameINPUT.TextAlign = "center"
                      $AddUserNameINPUT.Text = "User not found"
                 }
            } 
            else{return}
       }
        elseif($AddUserNameINPUT.Text -eq ""){
          $errorprovider.BlinkStyle ="NeverBlink"
          $errorprovider.SetIconPadding($AddUserNameINPUT, 1)
          $errorProvider.SetError($AddUserNameINPUT, "write User Name.")
        }#if user not found
        else{$Formherew.hide()
             $AddUserNameINPUT.BackColor ="red"
             $AddUserNameINPUT.TextAlign = "center"
             $AddUserNameINPUT.Text = "User not found"
        }
    }
 elseif($AddUserNameINPUT.Text -match $hebrewname){
       $adduserfilter = $AddUserNameINPUT.Text + "*"
       $SearchHebrewmail = SearchAD $adduserfilter
       $ResultPLus=@()
       foreach ($item in $SearchHebrewmail){
           $result=Get-ADUser $item |Select-Object name,SamAccountName
           $ResultPLus+=$result.name+" - "+$result.SamAccountName
       }
       if($SearchHebrewmail.Count -gt "1"){
           fromhebrew
          $AddUserNameINPUT.text=""
          $Split=$heblist.SelectedItem.Split("-")
          $Splituser=$Split[0].TrimStart(" ")
          $AddListMemebersImport.Items.Add($Splituser)
          return
       }
       elseif($SearchHebrewmail.Count -eq "1"){
               $hebrewnameaccount = get-aduser $SearchHebrewmail|Select-Object name
               $AddUserNameINPUT.text=""
               $AddListMemebersImport.Items.Add($hebrewnameaccount.name)
       }
       else{$Formherew.hide()
             $AddUserNameINPUT.BackColor ="red"
             $AddUserNameINPUT.TextAlign = "center"
             $AddUserNameINPUT.Text = "User not found"
       } 
    }
 elseif($AddUserNameINPUT.Text -eq ""){
          $errorprovider.BlinkStyle ="NeverBlink"
          $errorprovider.SetIconPadding($AddUserNameINPUT, 1)
          $errorProvider.SetError($AddUserNameINPUT, "write User Name.")}
 elseif($addname =Get-Mailbox $AddUserNameINPUT.Text |Select-Object name,SamAccountName){
           $AddListMemebersImport.Items.Add($addname.name)
           $AddUserNameINPUT.Text= ""
    }
 else{$Formherew.hide()
      $AddUserNameINPUT.BackColor ="red"
      $AddUserNameINPUT.TextAlign = "center"
      $AddUserNameINPUT.Text = "User not found"
 } 
}  
function closefrom(){
$Formherew.hide()
$AddListMemebersImport.Items.Clear()
}
function RemoveMembers(){
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
    foreach($AddItemsTOList in $msgListMemebers.SelectedItems){
            $MsgListRemove.Items.Add($AddItemsTOList)
    }
    foreach ($item in $MsgListRemove.Items){
        $msgListMemebers.Items.Remove($item)
    }
 }
 else{
   foreach($AddItemsTOList in $AddListMemebers.SelectedItems){
           $AddListRemove.Items.Add($AddItemsTOList) 
   }
   $AddNewList = $AddListMemebers.Items
   $Addlistnew = Compare-Object $AddNewList $AddListMemebers.SelectedItems | ForEach-Object { $_.InputObject }
   $AddListMemebers.Items.Clear()
   foreach($additems in $Addlistnew){
          $AddListMemebers.Items.Add($additems)
   }
 }     
}
function RemoveCalendar(){
 foreach ($usersremove in $AddListMemebersCalendar.SelectedItems){
          $AddListRemoveCalendar.Items.Add($usersremove)
 }
 $Addlistnew = Compare-Object $AddListMemebersCalendar.Items $AddListMemebersCalendar.SelectedItems | ForEach-Object { $_.InputObject }    
 $AddListMemebersCalendar.Items.Clear()
 foreach ($additem in $Addlistnew){
          $AddListMemebersCalendar.Items.Add($additem)
 }
}
function Importfile(){
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
   if($MsgListMemebersImport.Items.Count -gt "0"){
    $Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("This action will remove all users from Import members.`r`nDo you want to continue ? ",36, "Remove items")
  if($Questionremove -eq "yes"){
     $MsgListMemebersImport.Items.Clear()
     $MsgImportBtN.Enabled=$False
	 $MsgImportBtN.Text = "Loading..."
	 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
     $OpenFileDialog.Multiselect = $True
	 $OpenFileDialog.filter = "Text File (*.txt)| *.txt"
     $OpenFileDialog.ShowHelp = $True
	 $OpenFileDialog.ShowDialog() 
     $OpenFileDialog.DereferenceLinks
     $LoadFile = Get-Content $OpenFileDialog.FileNames -Encoding "Default"
     if($LoadFile){
        $names=@()
        foreach ($AddAccount in $LoadFile){
                 $MsgListMemebersImport.Items.Add($AddAccount)   
        }    
     }
     else{[Microsoft.VisualBasic.Interaction]::MsgBox("Sorry, no file was selected. Please select a file and try again.",16 , "No File Selected!")}
     $MsgImportBtN.Text = "Import File"
	 $MsgImportBtN.Enabled=$True
  }
  else{return}
 }
   else{$MsgListMemebersImport.Items.Clear()
        $MsgImportBtN.Enabled=$False
	    $MsgImportBtN.Text = "Loading..."
	    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Multiselect = $True
	    $OpenFileDialog.filter = "Text File (*.txt)| *.txt"
        $OpenFileDialog.ShowHelp = $True
	    $OpenFileDialog.ShowDialog() 
        $OpenFileDialog.DereferenceLinks
        $LoadFile = Get-Content $OpenFileDialog.FileNames -Encoding "Default"
     if($LoadFile){
        $names=@()
        foreach ($AddAccount in $LoadFile){
                 $MsgListMemebersImport.Items.Add($AddAccount)   
        }    
     }
     else{[Microsoft.VisualBasic.Interaction]::MsgBox("Sorry, no file was selected. Please select a file and try again.",16 , "No File Selected!")}
     $MsgImportBtN.Text = "Import File"
     $MsgImportBtN.Enabled=$True}
 }
 else{
  if($AddListMemebersImport.Items.Count -gt "0"){
    $Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("This action will remove all users from Import members.`r`nDo you want to continue ? ",36, "Remove items")
  if($Questionremove -eq "yes"){
     $AddListMemebersImport.Items.Clear()
     $AddImportBtN.Enabled=$False
	 $AddImportBtN.Text = "Loading..."
	 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
     $OpenFileDialog.Multiselect = $True
	 $OpenFileDialog.filter = "Text File (*.txt)| *.txt"
     $OpenFileDialog.ShowHelp = $True
	 $OpenFileDialog.ShowDialog() 
     $OpenFileDialog.DereferenceLinks
     $LoadFile = Get-Content $OpenFileDialog.FileNames -Encoding "Default"
     if($LoadFile){
        $names=@()
        foreach ($AddAccount in $LoadFile){
                 $AddListMemebersImport.Items.Add($AddAccount)   
        }    
     }
     else{[Microsoft.VisualBasic.Interaction]::MsgBox("Sorry, no file was selected. Please select a file and try again.",16 , "No File Selected!")}
     $AddImportBtN.Text = "Import File"
	 $AddImportBtN.Enabled=$True
  }
  else{return}
 }
  else{$AddListMemebersImport.Items.Clear()
     $AddImportBtN.Enabled=$False
	 $AddImportBtN.Text = "Loading..."
	 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
     $OpenFileDialog.Multiselect = $True
	 $OpenFileDialog.filter = "Text File (*.txt)| *.txt"
     $OpenFileDialog.ShowHelp = $True
	 $OpenFileDialog.ShowDialog() 
     $OpenFileDialog.DereferenceLinks
     $LoadFile = Get-Content $OpenFileDialog.FileNames -Encoding "Default"
     if($LoadFile){
        $names=@()
        foreach ($AddAccount in $LoadFile){
                 $AddListMemebersImport.Items.Add($AddAccount)   
        }    
     }
     else{[Microsoft.VisualBasic.Interaction]::MsgBox("Sorry, no file was selected. Please select a file and try again.",16 , "No File Selected!")}
  $AddImportBtN.Text = "Import File"
  $AddImportBtN.Enabled=$True}
 }
  
}
function ExportList(){
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
   $MsgExprotBTN.Enabled=$False
   $MsgExprotBTN.Text = "Loading..."
   if($ExportAcpet = (Get-Mailbox $MsgAliasINPUT.Text |Select-Object AcceptMessagesOnlyFromSendersOrMembers| Where {$_ -notlike <Domain>}).AcceptMessagesOnlyFromSendersOrMembers){
      $ListExport=@()
      foreach ($items in $ExportAcpet){
      if($List=Get-Mailbox "$items"|Select-Object DisplayName,SamAccountName,WindowsEmailAddress){$ListExport+=$List}
      elseif($List+=Get-DistributionGroup "$items"|Select-Object DisplayName,SamAccountName,WindowsEmailAddress){$ListExport+=$List}
      else{$ListExport+=Get-Contact "$items"|Select-Object DisplayName,SamAccountName,WindowsEmailAddress}
      }
      $ListExport|Out-GridView -Title "Accept Messages Only"|Sort-Object 
   }
   else{$ExportAcpet = (Get-DistributionGroup $MsgAliasINPUT.Text |Select-Object AcceptMessagesOnlyFromSendersOrMembers| Where {$_ -notlike <Domain>}).AcceptMessagesOnlyFromSendersOrMembers
        $ListExport=@()
        foreach ($items in $ExportAcpet){
        if($List=Get-Mailbox "$items"|Select-Object DisplayName,SamAccountName,WindowsEmailAddress){$ListExport+=$List}
        elseif($List+=Get-DistributionGroup "$items"|Select-Object DisplayName,SamAccountName,WindowsEmailAddress){$ListExport+=$List}
        else{$ListExport+=Get-Contact "$items"|Select-Object DisplayName,SamAccountName,WindowsEmailAddress}
        }
        $ListExport|Out-GridView -Title "Accept Messages Only"|Sort-Object  
   }
   $MsgExprotBTN.Text ="Exprot List"
   $MsgExprotBTN.Enabled=$True
}
 else{$AddExprotBTN.Enabled=$False
	  $AddExprotBTN.Text = "Loading..."
      if($AddFullRadio.Checked -eq $true){
         $AddMembersFull = Get-Mailbox $AddAliasINPUT.Text |get-MailboxPermission |Select-Object user,AccessRights | where { ($_.AccessRights -contains "FullAccess")}
         $addMembersName= $AddMembersFull.user |  where{($_ -ne <users>)}
         $AddMembers = $addMembersName.split("\") |  where{($_ -ne <users>)}
         $AddOnlyMembers= @()
         foreach($AddMembersfu in $AddMembers){
                 try{$AddOnlyMembers += get-ADUser $AddMembersfu -Properties * |Select-Object Name,uid|where{$_.uid -ne $null}}
                 catch{$error1= "all user"+$AddOnlyMembers}
         }
         $addNamelist = $AddOnlyMembers.name 
         $bresults=@()
         foreach ($memberlist in $addNamelist){
                  $userlist  = Get-Mailbox $memberlist|Select-Object SamAccountName
                  $bresults += Get-ADUser $userlist.SamAccountName -Properties * | Select-Object  DisplayName,SamAccountName,EmailAddress,`
                               Title,Uid,TelephoneNumber
         }
         $bresults| Sort-Object | Out-GridView -Title "full Access"
     }
      if($AddSendASRadio.Checked -eq $true){
         $addSenASMembers= Get-Mailbox $AddAliasINPUT.Text | Get-ADPermission | where {($_.ExtendedRights -like “*Send-As*”) -and -not ($_.User -like “NT AUTHORITY\SELF”)}|Select-Object user
         $addSenASMember = $addSenASMembers.user
         $addSenASSplit = $addSenASMember.split("\") |  where{$_ -ne <Domain>} 
         $addMemerSendasheb=@()
         foreach ($AddMembersSendas in $addSenASSplit){
                  $addMemerSendasheb += Get-ADUser $AddMembersSendas -Properties *|Select-Object DisplayName,SamAccountName,EmailAddress,`
                                        Title,Uid,TelephoneNumber
         }
         $addMemerSendasheb| Sort-Object | Out-GridView -Title "Send As"
      }
      if($AddSendOnRadio.Checked -eq $true){
         $AddsendOnMem = Get-Mailbox $AddAliasINPUT.Text |Select-Object GrantSendOnBehalfTo 
         $AddsendOnMemnot= $AddsendOnMem.GrantSendOnBehalfTo| Where {$_ -notlike <Domain>} 
         $AddSendoName =@()
         foreach($item in $AddsendOnMemnot){
                 $item2 = Get-Mailbox $item |Select-Object SamAccountName
                 $AddSendoName += Get-ADUser $item2.SamAccountName -Properties *|Select-Object DisplayName,SamAccountName,EmailAddress,`
                                  Title,Uid,TelephoneNumber
         }
         $AddSendoName| Sort-Object | Out-GridView -Title "send on behalf"
      }
      $AddExprotBTN.Text ="Exprot List"
	  $AddExprotBTN.Enabled=$True
 }
}
function MoveAccount(){
$errorprovider.Clear()
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
    if($MsgListMemebersImport.Items -gt "0"){
       if($MsgListMemebersImport.SelectedItems -gt "0"){
          $nousermail=@()
          foreach($AddAccountImport in $MsgListMemebersImport.SelectedItem){
                  $UserSerach=$AddAccountImport.TrimEnd(" ")
                  $usernamesinport = Get-Mailbox "$UserSerach"|Select-Object name
                  if($usernamesinport){$MsgListMemebers.Items.Add($usernamesinport.Name)
                                       $MsgListMemebersImport.Items.Remove($MsgListMemebersImport.SelectedItem)
                  }
                  elseif($UsersDistribution = (Get-DistributionGroup "$UserSerach"|Select-Object name).name){
                                              $MsgListMemebers.Items.Add($UsersDistribution)
                                              $MsgListMemebersImport.Items.Remove($MsgListMemebersImport.SelectedItem)
                  }
                  else{$nousermail += $AddAccountImport}
          }
          if($nousermail.Count -gt "0"){
             $usernone=@()
             For ($i=0; $i -lt $nousermail.Length -1 ; $i++){
                  $usernone +=$nousermail[$i] + ","}
                  $number=$nousermail.Length -1
                  $usernomail= "$usernone"+" " +$nousermail[$number]
                  [Microsoft.VisualBasic.Interaction]::MsgBox("לא נמצאו תיבות המייל, רשימות התפוצה שלהלן: $usernomail",16, "שגיאה")
          }
 }
       else{$errorprovider.BlinkStyle ="NeverBlink"
            $errorprovider.SetIconAlignment($MsgListMemebersImport, "MiddleLeft")
            $errorProvider.SetError($MsgListMemebersImport, "No account was selected, please selected.")
       }
    } 
 }
 elseif($AddListMemebersImport.Items -gt "0"){
    if($AddListMemebersImport.SelectedItems -gt "0"){
       $nousermail=@()
       foreach($AddAccountImport in $AddListMemebersImport.SelectedItems){
               $UserSerach=$AddAccountImport.TrimEnd(" ")
               $usernamesinport = Get-Mailbox "$UserSerach"|Select-Object name
            if($usernamesinport){$AddListMemebers.Items.Add($usernamesinport.Name)}
            else{$nousermail += $AddAccountImport}
       }
       if($nousermail.Count -gt "0"){
                  $usernone=@()
                    For ($i=0; $i -lt $nousermail.Length -1 ; $i++)
                         {$usernone +=$nousermail[$i] + ","}
                          $number=$nousermail.Length -1
                          $usernomail= "$usernone"+" " +$nousermail[$number]
                           [Microsoft.VisualBasic.Interaction]::MsgBox("לא נמצאו תיבות המייל שלהלן: $usernomail",16, "שגיאה")
                        } 
    $AddAccountList = $AddListMemebersImport.Items
    $AddAccountClear = (Compare-Object $AddAccountList $AddListMemebersImport.SelectedItems|?{($_.SideIndicator -eq "=>")}).InputObject
    if($AddAccountClear -eq $null){
       $AddListMemebersImport.Items.add($AddAccountClear)
       $AddListMemebersImport.Items.Remove($AddAccountClear)
       foreach($addAccountTolist in $AddAccountClear){
             $AddListMemebersImport.Items.Add($addAccountTolist)}
   }
    }
    else{[void][System.Windows.Forms.MessageBox]::Show("The user $AddAccountClear already has access to this mailbox.","Shared MailBox")}

   else{$errorprovider.BlinkStyle ="NeverBlink"
         $errorprovider.SetIconAlignment($AddListMemebersImport, "MiddleLeft")
         $errorProvider.SetError($AddListMemebersImport, "No account was selected, please selected.")
    }
 }
}
function MoveAll(){
 $errorprovider.Clear()
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
    $usernamesinport=@()
    $nousermail=@()
    foreach($AddAccountImport in $MsgListMemebersImport.Items){
            $User= $AddAccountImport.TrimEnd("")
            $mailusers = Get-Mailbox "$User"|Select-Object name
          if($mailusers){$usernamesinport += $mailusers}
          elseif(($UsersDistribution= Get-DistributionGroup "$User"|Select-Object name).name){$usernamesinport += $UsersDistribution}
          else{$nousermail += $AddAccountImport}
    }
    if($nousermail.Count -gt "0"){
       $usernone=@()
       For ($i=0; $i -lt $nousermail.Length -1 ; $i++){
       $usernone +=$nousermail[$i] + ","}
       $number=$nousermail.Length -1
       $usernomail= "$usernone"+" " +$nousermail[$number]
       [Microsoft.VisualBasic.Interaction]::MsgBox("לא נמצאו תיבות המייל, רשימות התפוצה שלהלן: $usernomail",16, "שגיאה")
    }
    
    foreach($AddListMemebersconter in $usernamesinport.name){
           $MsgListMemebers.Items.Add($AddListMemebersconter)
    }
           $MsgListMemebersImport.Items.Clear()
    }
 Elseif($Tabctrl.SelectedTab -eq $AddTAB){
  $usernamesinport=@()
  $nousermail=@()
  foreach($AddAccountImport in $AddListMemebersImport.Items){
          $User= $AddAccountImport.TrimEnd("")
     $mailusers = Get-Mailbox "$User"|Select-Object name
     if($mailusers){$usernamesinport += Get-Mailbox "$User"|Select-Object name}
     else{$nousermail += $AddAccountImport}
  }
  if($nousermail.Count -gt "0"){
     $usernone=@()
     For ($i=0; $i -lt $nousermail.Length -1 ; $i++){
     $usernone +=$nousermail[$i] + ","}
     $number=$nousermail.Length -1
     $usernomail= "$usernone"+" " +$nousermail[$number]
     [Microsoft.VisualBasic.Interaction]::MsgBox("לא נמצאו תיבות המייל שלהלן: $usernomail",16, "שגיאה")
  }
  $MoveAllList=(Compare-Object $usernamesinport.name $AddListMemebers.Items|?{($_.SideIndicator -eq "<=")}).InputObject
  foreach($AddListMemebersconter in $MoveAllList){
          $AddListMemebers.Items.Add($AddListMemebersconter)}
          $AddListMemebersImport.Items.Clear()
 }
}
function SaveAll(){
 if($Tabctrl.SelectedTab -eq $AcceptMsgTAB){
    if($MsgListRemove.Items -gt "0"){
       $Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("You are about to remove permission, are you sure?",36, "Remove users")
       if($Questionremove -eq "yes"){
         $Users=$MsgListRemove.Items
         $MemberList=(Get-Mailbox $MsgAliasINPUT.text |Select-Object AcceptMessagesOnlyFromSendersOrMembers).AcceptMessagesOnlyFromSendersOrMembers
         $ListNot=$MemberList|?{$_ -notlike "*$Users"}
         Set-Mailbox $MsgAliasINPUT.text -AcceptMessagesOnlyFromSendersOrMembers $ListNot
       }
    }
    elseif($MsgListMemebers.Items -gt "0"){
    
    }
 }
 elseif($AddFullRadio.Checked -eq $true){
 if($AddListRemove.Items -gt "0"){
 $Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("You are about to remove permission from the mailbox, are you sure?",36, "Remove users")
 if($Questionremove -eq "yes"){
  foreach($AddRemoveusr in $AddListRemove.Items){
  $usersam=Get-Mailbox "$AddRemoveusr"|Select-Object SamAccountName
   $AddRemoveusr = remove-MailboxPermission $AddAliasINPUT.Text -User $usersam.SamAccountName -AccessRights FullAccess  -Confirm:$false}
    $AddListRemove.Items.Clear()
     if(!($AddListMemebers.Items)){
     ClearForm3
     [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions"+ " Full access" + " have been removed for the selected users.",64, "Success")
     <#סוף מחיקה#>}}
     else{}
     }
 if($AddListMemebers.Items -gt "0"){
     $addGETUsers = Get-Mailbox $AddAliasINPUT.Text |get-MailboxPermission |Select-Object user,AccessRights | where { ($_.AccessRights -contains "FullAccess")} 
     $AddMemebersc = $addGETUsers.user
     $AddMemebersplit =  $AddMemebersc.split("\") |  where{($_ -ne <Domain>)}
       $fullperusersmission=@()
        foreach($fulluser in $AddMemebersplit){
                try{$fullperusersmission += get-ADUser $fulluser -Properties * |Select-Object Name,uid,SamAccountName|where{$_.uid -ne $null}}
                 catch{$error3= "all user"+$fullperusersmission}}
   if($fullperusersmission -ne $null){
       $AddCompareuser = Compare-Object $fullperusersmission.name $AddListMemebers.Items| ForEach-Object { $_.InputObject }
         foreach ($item in $AddCompareuser){
                  $Addfullpermission = add-MailboxPermission $AddAliasINPUT.Text -user $item -AccessRights FullAccess}
                  ClearForm3
                   [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + " have been granted for the selected users.",64, "Success") }                  
   else{foreach($addmemebersone in $AddListMemebers.Items){
   $UserSamAccountName=Get-Mailbox "$addmemebersone" |Select-Object SamAccountName
                 $Addfullpermission = add-MailboxPermission $AddAliasINPUT.Text -user $UserSamAccountName.SamAccountName -AccessRights FullAccess}
                 ClearForm3
                 [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + " have been granted for the selected users.",64, "Success")
 }
                 }
}
 elseif($AddSendASRadio.Checked -eq $true){
 if($AddListRemove.Items -gt "0"){
    $Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("You are about to remove permission from the mailbox, are you sure?",36, "Remove users")
 if($Questionremove -eq "yes"){
    $user=@()
    $sam = Get-Mailbox $AddAliasINPUT.Text|Select-Object SamAccountName
    $Distinname = Get-ADUser $sam.SamAccountName |Select-Object DistinguishedName
    $user=@()
     foreach ($item2 in $AddListRemove.Items){
              $user += Get-Mailbox "$item2"|Select-Object SamAccountName}
     foreach ($AddRemoveusr1 in $user.SamAccountName){
              $AddRemoveusrsendas= Remove-ADPermission -Identity $Distinname.DistinguishedName -User "<Domain>\$AddRemoveusr1" `
              -InheritanceType 'All' -ExtendedRights 'send-as' -confirm:$false
              $AddRemovefull= remove-MailboxPermission $AddAliasINPUT.Text -User $AddRemoveusr1 -AccessRights FullAccess  -Confirm:$false}
   if(!($AddListMemebers.Items)){
    ClearForm3
    [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + "&" + " Send as" + "have been removed for the selected users.",64, "Success")
     <#סוף הסרה#>} }}                    
 if($AddListMemebers.Items -gt "0"){
     $addGETUsers = Get-Mailbox $AddAliasINPUT.Text | Get-ADPermission | where {($_.ExtendedRights -like “*Send-As*”)}|Select-Object user 
     $AddMemebersc = $addGETUsers.user
     $AddMemebersplit = $AddMemebersc.split("\") |  where{($_ -ne <Domain>)}
      $fullperusersmission=@()
      foreach($fulluser in $AddMemebersplit){
                try{$fullperusersmission += get-ADUser $fulluser -Properties * |Select-Object Name,uid,SamAccountName|where{$_.uid -ne $null}}
                 catch{$error3= "all user"+$fullperusersmission}}
   if($fullperusersmission -ne $null){
       $AddCompareuser = Compare-Object $fullperusersmission.name $AddListMemebers.Items |Where{$_.SideIndicator -eq "=>"}|Select-Object InputObject
        $samm=@()
        foreach ($sumsum in $AddCompareuser.InputObject){
                 $samm += Get-Mailbox $sumsum |Select-Object SamAccountName
             }
         foreach($item in $samm.SamAccountName){
                  $Addfullpermission = add-MailboxPermission $AddAliasINPUT.Text -user $item -AccessRights FullAccess
                  $AddSendApermission = Get-Mailbox $AddAliasINPUT.Text | add-ADPermission -User $item -ExtendedRights "Send As"}      
                  ClearForm3
                   [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + "&" + " Send as" + " have been granted for the selected users.",64, "Success")} 
   else{$samm=@()
        foreach ($sumsum in $AddListMemebers.Items){
                 $samm += Get-Mailbox $sumsum |Select-Object SamAccountName
                                                    }
        foreach($addmemebersone in $samm.SamAccountName){
                     $Addfullpermission = add-MailboxPermission $AddAliasINPUT.Text -user $addmemebersone -AccessRights FullAccess
                     $AddSendApermission = Get-Mailbox  $AddAliasINPUT.Text | add-ADPermission -User $addmemebersone -ExtendedRights "Send As"}
                     ClearForm3
                     [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + "&" + " Send as" + " have been granted for the selected users.",64, "Success")}
}
}
 elseif($AddSendOnRadio.Checked -eq $true){
 if($AddListRemove.Items -gt "0"){
   <#הסרה#>$Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("You are about to remove permission from the mailbox, are you sure?",36, "Remove users")
 if($Questionremove -eq "yes"){
  $addsendonusers = get-mailbox $AddAliasINPUT.Text | Select-Object grantsendonbehalfto
    $addsenonremoveold = $addsendonusers.GrantSendOnBehalfTo|where{$_ -notlike <domain>}
     $addremovesendonper =@()
     foreach($removesendon in $AddListRemove.Items){
             $addremovesendonper += Get-Mailbox $removesendon |Select-Object SamAccountName,Identity}
             $addremovesenonwhere = $addremovesendonper.Identity
             $addSendREmove = $addremovesendonper.SamAccountName
 $addsenonlistfin = Compare-Object $addsenonremoveold $addremovesendonper.Identity |Where{$_.SideIndicator -eq "<="}|Select-Object InputObject
       set-mailbox -Identity $AddAliasINPUT.Text -GrantSendOnBehalfTo $addsenonlistfin.InputObject
      foreach($addremovesend in $addSendREmove){
       $AddRemovesendfull= remove-MailboxPermission $AddAliasINPUT.Text -User $addremovesend -AccessRights FullAccess  -Confirm:$false}
        if(!($AddListMemebers.Items)){
        ClearForm3
        [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + "&" + " Send on behalf" + " have been removed for the selected users.",64, "Success")} <#סוף הסרה}#>}
       else{} }
 if($AddListMemebers.Items -gt "0"){
      $addsendonuserss = get-mailbox $AddAliasINPUT.Text | Select-Object grantsendonbehalfto
    $addsendonwithoutold = $addsendonuserss.grantsendonbehalfto | where{$_ -notlike <domain>}
    $addlistsendon=@()
    foreach ($Addsendon in $AddListMemebers.Items){
             $addlistsendon += Get-Mailbox $Addsendon |Select-Object Identity,SamAccountName}
    $addlistsendonbe = @()
    $addlistsendonbe +=  $addlistsendon.Identity
        set-mailbox -Identity $AddAliasINPUT.Text -GrantSendOnBehalfTo $addlistsendonbe
     foreach ($addacount in $addlistsendon.SamAccountName ){
          $AddSendpermission = add-MailboxPermission $AddAliasINPUT.Text -user $addacount -AccessRights FullAccess  
        }
        ClearForm3
        [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " Full access" + "&" + " Send on behalf" + " have been granted for the selected users.",64, "Success")}
}
 elseif($AddcalendarRadio.Checked -eq $true){
    if($AddListRemoveCalendar.Items -gt "0"){
    $Questionremove = [Microsoft.VisualBasic.Interaction]::MsgBox("You are about to remove permission from the calendar, are you sure?",36, "Remove users")
    if($Questionremove -eq "yes"){
        $itemtoremove = $AddListRemoveCalendar.Items
        $aliasname = $AddAliasINPUT.Text
        foreach ($item in $AddListRemoveCalendar.Items){
                  $ItemSplit=$item.split("-")
                  if(($ItemSplit[2] -eq " Reviewer") -or ($ItemSplit[1] -eq " Reviewer")){
                      $items=$item.TrimEnd(" - Reviewer")
                  }
                  elseif(($ItemSplit[2] -eq " Editor") -or ($ItemSplit[1] -eq " Editor")){
                          $items=$item.TrimEnd(" - Editor")
                  }
                                  
                  <#
                  $itemsplit=$item.split("-")
                  $items = $itemsplit[0] # |where{($_ -ne  " Reviewer") -or ($_ -ne  " Editor")}
                  #>
                  $samitem =Get-Mailbox $items | Select-Object SamAccountName
                    $RemoveUserPermission =remove-MailboxFolderPermission -Identity $aliasname":\לוח שנה" -User  $samitem.SamAccountName -Confirm:$false
                    if(!($RemoveUserPermission)){
                    $RemoveUserPermission2=remove-MailboxFolderPermission -Identity $aliasname":\calendar" -User  $samitem.SamAccountName -Confirm:$false}
                             }                                               
          if(!($AddListMemebersCalendarImport.Items -gt "0")){ 
          ClearForm3
          [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " have been removed for the selected users.",64, "Success")} <#סוף הסרה}#>}
    }                               
    if($AddListMemebersCalendarImport.Items -gt "0"){
         foreach ($names in $AddListMemebersCalendarImport.Items){
                  $account=($names.split("<"))[$_.count-1].Trim("`n- ")
                  $permission=($names.split("<"))[$_.count-2].Trim("`n- ")
                  $Useraccount =(Get-Mailbox "$account" |Select-Object SamAccountName).SamAccountName
                  if($permission -match "Reviewer"){
                     $mailheb=add-MailboxFolderPermission -Identity ("{0}:\לוח שנה"-f $AddAliasINPUT.Text) -User $Useraccount -AccessRights Reviewer
                     if(!($mailheb)){add-MailboxFolderPermission -Identity ("{0}:\calendar"-f $AddAliasINPUT.Text) -User $Useraccount -AccessRights Reviewer}
                  }
                  if($permission -match "Editor"){
                     $mailheb =add-MailboxFolderPermission -Identity ("{0}:\לוח שנה"-f $AddAliasINPUT.Text)  -User $Useraccount -AccessRights Editor
                     if(!($mailheb)){add-MailboxFolderPermission -Identity ("{0}:\calendar"-f $AddAliasINPUT.Text) -User $Useraccount -AccessRights Editor}
                  }
         }
     ClearForm3                                          
     [Microsoft.VisualBasic.Interaction]::MsgBox("The Permissions" + " have been granted for the selected users.",64, "Success")
     }
 }
}
#---------------------------------------function Msg Tab-------------------------------------#
function CheckAccpet(){
 if(($MsgAliasINPUT.Text -notlike "Mail not found for*")-and($MsgAliasINPUT.text -ne "Alias Name Only")-and($MsgAliasINPUT.text -ne "Plase fill an alias name")`
     -and($MsgAliasINPUT.text -ne "Fax Number Not Found")){
  if($MsgAliasINPUT.Text -match $pat){
    $faxother = "*"+$MsgAliasINPUT.Text +"*"
    $faxuser=(Get-ADUser -filter {otherFacsimileTelephoneNumber -like $faxother}|Select-Object SamAccountName).SamAccountName
    if($faxuser -gt "0"){
       if($faxuser.SamAccountName.count -eq "1"){$AddAliasINPUT.Text = $faxuser.SamAccountName}
       else{[Microsoft.VisualBasic.Interaction]::MsgBox("More than one user was found.`n`a Please write down the user's Alias name.",16, "Error")
            return
       }
    }
    else{$MsgAliasINPUT.TextAlign = "center"
         $MsgAliasINPUT.BackColor ="red"
         $MsgAliasINPUT.Text = "Fax Number Not Found"
         return
   }                              
 }
  if($MsgAliasINPUT.Text -match $hebrewname){
     $hebrewmail = $MsgAliasINPUT.Text + "*"
     $SearchHebrewmail = SearchHeb $hebrewmail| Sort-Object -Unique
     if($SearchHebrewmail.Count -gt "1"){
          $ResultPLus=@()
          foreach ($item in $SearchHebrewmail){
           try{$Result=Get-ADUser $item |Select-Object name,SamAccountName
               $ResultPLus+=$result.name #+" - "+$result.SamAccountName
           }
           catch{$Result=Get-DistributionGroup "$item"|Select-Object name,SamAccountName
                 $ResultPLus+=$result.name #+" - "+$result.SamAccountName
           }
          }
          aliasfrom
          if (!($aliasLST.SelectedItem.count -eq "0")){
                #$Split=$aliasLST.SelectedItem.Split("-")
                #$Splituser=$Split[1].TrimStart(" ")
                if($MsgAliasINPUT.Text=(Get-Mailbox $aliasLST.SelectedItem|Select-Object PrimarySmtpAddress).PrimarySmtpAddress){
                   $MsgAliasINPUT.Enabled = $false
                   $MsgAliasINPUT.textalign="center"
                }
                elseif($MsgAliasINPUT.Text=(Get-DistributionGroup $aliasLST.SelectedItem|Select-Object PrimarySmtpAddress).PrimarySmtpAddress){
                       $MsgAliasINPUT.Enabled = $false
                       $MsgAliasINPUT.textalign="center"
                }
                else{$MsgAliasINPUT.TextAlign = "center"
                     $MsgAliasINPUT.BackColor ="red"
                     $MsgAliasINPUT.Text = "Mail not found for $Splituser"
                     return
                 }           
       }
       }
     elseif($SearchHebrewmail.Count -eq "1"){
              $Addlist= (Get-Mailbox $SearchHebrewmail|Select-Object PrimarySmtpAddress).PrimarySmtpAddress
              if($Addlist){
                 $MsgAliasINPUT.Enabled = $false
                 $MsgAliasINPUT.textalign="center"
                 $MsgAliasINPUT.Text=$Addlist
                 
              }
              elseif($Addlist= (Get-DistributionGroup $SearchHebrewmail|Select-Object PrimarySmtpAddress).PrimarySmtpAddress){
                     $MsgAliasINPUT.Enabled = $false
                     $MsgAliasINPUT.textalign="center"
                     $MsgAliasINPUT.Text=$Addlist
                     
              }
              else{$Formherew.hide()
                   $MsgAliasINPUT.TextAlign = "center"
                   $MsgAliasINPUT.BackColor ="red"
                   $user=$MsgAliasINPUT.text
                   $MsgAliasINPUT.Text = "Mail not found for $user"
                   return
            }
       }
     else{$Formherew.hide()
          $MsgAliasINPUT.TextAlign = "center"
          $MsgAliasINPUT.BackColor ="red"
          $user=$MsgAliasINPUT.text
          $MsgAliasINPUT.Text = "Mail not found for $user"
            return
       }
}
  If($MsgAliasINPUT.Text -match $EnglishName){
     $MsgChackMail = Get-Mailbox $MsgAliasINPUT.Text|Select-Object WindowsEmailAddress,AcceptMessagesOnlyFromSendersOrMembers
     if($MsgChackMail -ne $null){
        $MsgAliasINPUT.text= $MsgChackMail.WindowsEmailAddress
        $MsgAliasINPUT.Enabled = $false
        $MsgAliasINPUT.TextAlign = "center"
        $AcceptMsgList=$MsgChackMail.AcceptMessagesOnlyFromSendersOrMembers
        if($AcceptMsgList.count -gt "0"){
           foreach($addshowcontrols in $MsgTbObject){
                   $AcceptMsgTAB.Controls.Add($addshowcontrols)
           }
           foreach ($list in $AcceptMsgList){
                    $UserName= Split-Path  $list -leaf
                    $MsgListMemebers.items.add($UserName)        
           }
        }
        else{$AddMsgBox =[Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this mailbox, Do you wish to continue?",36, "Error Permission")
             if($AddMsgBox -eq "yes"){ 
                foreach($addshowcontrols in $MsgTbObject){
                        $AcceptMsgTAB.Controls.Add($addshowcontrols)
                }
                $MsgAliasINPUT.Enabled = $false
             }
             else{ClearAccpetTab}
            }
     }
     elseif($MsgChackMail= Get-DistributionGroup $MsgAliasINPUT.Text|Select-Object WindowsEmailAddress,AcceptMessagesOnlyFromSendersOrMembers){
            $MsgAliasINPUT.text= $MsgChackMail.WindowsEmailAddress
            $MsgAliasINPUT.Enabled = $false
            $MsgAliasINPUT.TextAlign = "center"
            $AcceptMsgList=$MsgChackMail.AcceptMessagesOnlyFromSendersOrMembers
            if($AcceptMsgList.count -gt "0"){
               foreach($addshowcontrols in $MsgTbObject){
                       $AcceptMsgTAB.Controls.Add($addshowcontrols)
               }
               foreach ($list in $AcceptMsgList){
                        $UserName= Split-Path  $list -leaf
                        $MsgListMemebers.items.add($UserName)        
               }
            }
            else{$AddMsgBox =[Microsoft.VisualBasic.Interaction]::MsgBox("No one has permission to this mailbox, Do you wish to continue?",36, "Error Permission")
                    if($AddMsgBox -eq "yes"){ 
                       foreach($addshowcontrols in $MsgTbObject){
                               $AcceptMsgTAB.Controls.Add($addshowcontrols)
                       }
                       $MsgAliasINPUT.Enabled = $false
                    }
                    else{ClearAccpetTab}
            }
     }
     else{$MsgAliasINPUT.TextAlign = "center"
          $MsgAliasINPUT.BackColor ="red"
          $user=$MsgAliasINPUT.text
          $MsgAliasINPUT.Text = "Mail not found for $user"
     }
  }
  else{$MsgAliasINPUT.TextAlign = "center"
      $MsgAliasINPUT.BackColor ="red"
      $MsgAliasINPUT.Text = "Plase fill an alias name"
  }
 }
} 
function SearchHeb($Alias){
    $result = @()
    $distributiongroup=@()
    $result += Get-ADUser  -Filter {Name -like $Alias}| ? { ($_.distinguishedname -notlike '*<Domain>*') }|Select-Object name,samaccountname
    $resultsam = Get-ADUser -Filter {SamAccountName -eq $Alias }
    $resultdispaly=Get-ADUser -Properties DisplayName -Filter{DisplayName -like $Alias}| ? { ($_.distinguishedname -notlike '<Domain>') }|Select-Object name,samaccountname
    $distributiongroup=Get-DistributionGroup -filter "DisplayName -like '$alias'" |Select-Object name,samaccountname
    if(($result.samaccountname -gt "1")-and($resultdispaly.samaccountname -gt "1")-and($distributiongroup.samaccountname -gt "1")){
           $result=$result.samaccountname;$resultdispaly.samaccountname;$distributiongroup.samaccountname|Sort-Object -Unique
           return $a
    }
    elseif($result.samaccountname -gt "1"){
           return $result.SamAccountName
    }
    elseif($resultdispaly.samaccountname -gt "1"){
           return $resultdispaly.SamAccountName
    }
    elseif($distributiongroup.samaccountname -gt "1"){
           return $distributiongroup.SamAccountName
    }
}
function ClearAccpetTab(){
$MsgAliasINPUT.TextAlign="center"
    $MsgAliasINPUT.Text="Alias Name Only"
    $MsgAliasINPUT.Enabled=$True
    $MsgListMemebers.Items.Clear()
    $MsgListMemebersImport.items.clear()
    $MsgListRemove.items.clear()
    foreach($addshowcontrols in $MsgTbObject){
            $AcceptMsgTAB.Controls.remove($addshowcontrols)
    }
}
#--------------------------------------Object Main form--------------------------------------#
$FormMain = GenerateForm -title 'Shared MailBox | Version 1.11' -Width 420 -Height 470
 $FormMain.minimumSize = New-Object System.Drawing.Size(420,480)
 $FormMain.maximumSize = New-Object System.Drawing.Size(420,480)
 $FormMain.KeyPreview = $True
 $FormMain.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$FormMain.Close()}})
 #$FormMain.Add_KeyDown({if ($_.KeyCode -eq "Enter"){if($Tabctrl.SelectedTab -eq $AddTAB -and $AddSendOnRadio.Checked -eq $true`
 #-or $AddcalendarRadio.Checked -eq $true -or $AddSendASRadio.Checked -eq $true -or $AddFullRadio.Checked -eq $true){addUsrMail}}}) 
#-------------------------------------------tabs-------------------------------------------#
$NewTAB = Gentabpage -text 'New Shared MailBox'
$AddTAB = Gentabpage -text 'Add Members To Mail'
$AcceptMsgTAB = Gentabpage -text 'Accept messages from'
$Tabctrl = Gentabctrl -width 400 -height 440 -x 3 -y 0 -tabpage $NewTAB,$AddTAB <#,$AcceptMsgTAB#>
$Tabctrl.SelectedTab = $NewTAB <#$AcceptMsgTAB#>
#-------------------------------------------Calendar form-------------------------------------------#
$CalendarMain = GenerateForm -title 'Calendar permissions' -Width 260 -Height 150
 $CalendarMain.minimumSize = New-Object System.Drawing.Size(260,150)
 $CalendarMain.maximumSize = New-Object System.Drawing.Size(260,150)
$choosepermissionsLABEL = GenerateLabel -text 'Choose permission:' -x 45 -y 1
 $choosepermissionsLABEL.TextAlign = "MiddleLeft"
 $choosepermissionsLABEL.Font= New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)

$choosepermissionslist = GenerateListBox -x 77 -y 35  -width 80 -height 40
$choosepermissionslist.Add_Click{if($choosepermissionslist.SelectedItem.count -eq "1"){$CalendarMain.Hide()}
                                 else{[Microsoft.VisualBasic.Interaction]::MsgBox("No type permission was selected, Please select",16, "Error")}
                                 }
$add = $choosepermissionslist.Items.Add("Editor")|Out-Null
$add1 = $choosepermissionslist.Items.Add("Reviewer")|Out-Null

$choosepermissionsbtn= GenerateButton -text 'Choose' -x 120 -y 75 -action{if($choosepermissionslist.SelectedItem.count -eq "1"){$CalendarMain.close()}
                                                                          else{[Microsoft.VisualBasic.Interaction]::`
                                                                          MsgBox("No type permission was selected, Please select",16, "Error")}
                                                                          }
$choosepermissionscentelbtn= GenerateButton -text 'Cancel' -x 40 -y 75 -action{$CalendarMain.close()}

$AddListmemberCalendarLABEL = GenerateLabel -text 'Current members:' -x 40 -y 130
$AddListMemebersCalendar =  GenerateListBox -x 7 -y 145 -width 160 -height 214
$AddListMemebersCalendar.RightToLeft ="yes"
$AddListMemebersCalendar.MultiColumn = "true"
$AddListMemebersCalendar.SelectionMode = "MultiSimple"
$AddremoveCalendar = GenerateButton -text 'Remove' -x 60 -y 360 -action {RemoveCalendar}

$AddListMemebersCalendarImportLABEL = GenerateLabel -text 'New Members:' -x 260 -y 130
$AddListMemebersCalendarImport = GenerateListBox -x 215 -y 145 -width 165 -height 100
$AddListMemebersCalendarImport.RightToLeft ="yes"

$AddListRemoveCalendarLABEL = GenerateLabel -text 'Remove Memebers:' -x 250 -y 245
$AddListRemoveCalendar = GenerateListBox -x 215 -y 260 -width 165 -height 100
$AddListRemoveCalendar.RightToLeft ="yes"
#-------------------------------------------formHeb-------------------------------------------#
$Formherew = GenerateForm -title 'Choose Hebrew Name' -Width 250 -Height 250
 $Formherew.minimumSize = New-Object System.Drawing.Size(250,250)
 $Formherew.maximumSize = New-Object System.Drawing.Size(250,250)
 $Formherew.KeyPreview = $True
 $Formherew.Add_KeyDown({if ($_.KeyCode -eq "Enter"){ChooseName}}) 
 $Formherew.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Formherew.hide()}})
# $Formherew.add_Closing({echo "123"|Out-Null})

$hebChooseUserNameLABEL = GenerateLabel -text 'Choose UserName' -x 40 -y 8
 $hebChooseUserNameLABEL.TextAlign = "MiddleLeft"
 $hebChooseUserNameLABEL.Font= New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)

$heblist = GenerateListBox -x 40 -y 34 -width 150 -height 150
$heblist.SelectionMode = "MultiSimple"
$heblist.RightToLeft = "yes"
$heblist.Add_Click{ChooseName}
$hebchoosebtn = GenerateButton -text 'Choose'-x 40 -y 185 -action{ChooseName}
$hebcancelbtn =GenerateButton -text 'Cancel' -x 117 -y 185 -action {$Formherew.hide()
 $AddUserNameINPUT.Text= ""}
#-------------------------------------------Alias hebrew-------------------------------------------#
$Formaliashebrew = GenerateForm -title 'Choose Hebrew Alias' -Width 250 -Height 250
 $Formaliashebrew.minimumSize = New-Object System.Drawing.Size(250,250)
 $Formaliashebrew.maximumSize = New-Object System.Drawing.Size(250,250)
 $Formaliashebrew.KeyPreview = $True
 $Formaliashebrew.Add_KeyDown({if ($_.KeyCode -eq "Enter"){ChooseName}}) 
 $Formaliashebrew.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Formherew.hide()}})
 $Formaliashebrew.Add_FormClosing({if($aliasLST.SelectedItem -eq $null){
                                      $Formaliashebrew.Hide();ClearForm}
                                   else{} }) 

$choosenameLABEL = GenerateLabel -text 'Choose name' -x 40 -y 8
 $choosenameLABEL.TextAlign = "MiddleLeft"
 $choosenameLABEL.Font= New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)

$aliasLST = GenerateListBox -x 40 -y 34 -width 150 -height 150
$aliasLST.RightToLeft = "yes"

$aliasLST.Add_Click{ChooseName}

$aliaschoosebtn = GenerateButton -text 'Choose'-x 40 -y 185 -action{ChooseName}
$aliascancelbtn =GenerateButton -text 'Cancel' -x 117 -y 185 -action {$Formaliashebrew.Hide();ClearForm}
#-------------------------------------------Newtab-------------------------------------------#
$NewNameLABEL = GenerateLabel -text 'Name:' -x 5 -y 20
 $NewNameLABEL.Font= New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)
$NewNameINPUT = GenTextBox -x 65 -y 20 -width 200
 $NewNameINPUT.TabIndex = "1"
 $NewNameINPUT.RightToLeft = "yes"
 $NewNameINPUT.Add_Click({
 $NewNameINPUT.TextAlign = "left"
 $NewNameINPUT.BackColor = "window"})

$NewNamecheck = GenerateButton -text 'Check Available' -x 285 -y 18 -width 95 -height 25 -action {CheckAvailable}
 $NewNamecheck.Size = New-Object System.Drawing.Size(95, 25)
 $NewNamecheck.TabIndex = "2"
$NewAliasLABEL = GenerateLabel -text 'Alias:' -x 5 -y 60
 $NewAliasLABEL.Font = New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)

$NewAliasINPUT = GenTextBox -x 65 -y 60 -width 200 -height 100
 $NewAliasINPUT.TabIndex = "3"
 $NewAliasINPUT.Add_Click({
 $NewAliasINPUT.TextAlign = "left"
 $NewAliasINPUT.BackColor = "window"})

$NewAliascheck = GenerateButton -text 'Check Available' -x 285 -y 58 -width 95 -height 25 -action {CheckAvailable2}
 $NewAliascheck.Size = New-Object System.Drawing.Size(95, 25)
 $NewAliascheck.TabIndex = "4"

$NewDbLABEL = GenerateLabel -text 'DB:' -x 5 -y 100
 $NewDbLABEL.Font = New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)

 $newDbData = Get-MailboxDatabase|Select-Object name|Sort-Object name

$NewDbcombobox = GenComboBox -x 65 -y 100 -width 110 -height 50 -data $newDbData.name
 $NewDbcombobox.DropDownStyle ="dropdownlist"

$NewGropBox = GenGroupBox -width 385  -height 140 -x 2 -y 3

$newWhomanageLABEL = GenerateLabel -text 'Description:' -x 5 -y 160
 $newWhomanageLABEL.Font = New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)
$newWhomanageINPUT  = GenTextBox -x 112 -y 160 -width 240
 $newWhomanageINPUT.RightToLeft = "yes"
 $newWhomanageINPUT.Add_Click({
 $newWhomanageINPUT.TextAlign = "left"
 $newWhomanageINPUT.BackColor = "window"})
 $newWhomanageINPUT.TabIndex = "6"
$NewGropBox2 = GenGroupBox -width 385 -height 50 -x 2 -y 145 

$NewclearBtn = GenerateButton -text 'Clear' -x 40 -y 390 -action {clearNewMail}
$NewCreateBtn = GenerateButton -text 'Create' -x 150 -y 390 -action {CreateMail}
$NewCencal = GenerateButton -text 'Exit' -x 260 -y 390 -action {if([System.Windows.Forms.MessageBox]::Show("Do you wish to close the program?",` 
        "Shared MailBox",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "no"){
        return
    }
   $FormMain.Close()}
#-------------------------------------------Addtab-------------------------------------------#
$AddAliasLABEL = GenerateLabel -text 'Alias:' -x 5 -y 20
$AddAliasLABEL.Font= New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)
$AddAliasINPUT = GenTextBox -x 60 -y 20 -width 315
$AddAliasINPUT.text = "Alias Name Only"
$AddAliasINPUT.TextAlign ="center"
$AddAliasINPUT.Add_Click({
$AddAliasINPUT.TextAlign = "left"
$AddAliasINPUT.Text = ""
$AddAliasINPUT.BackColor = "window"})
 $AddAliasINPUT.TabIndex = "1"
$AddFullRadio = GenRadioBox -x 52 -y 45 -width 80 -height 15 -text 'FA'
 $AddFullRadio.TabIndex = "2"
 $AddFullRadio.Add_CheckedChanged{
  if($AddFullRadio.Checked -eq $true){FullAcess}}
$AddSendASRadio = GenRadioBox -x 92 -y 45 -width 93 -height 15 -text 'Send As + FA'
$AddSendASRadio.TabIndex = "3"
 $AddSendASRadio.Add_CheckedChanged{
 if($AddSendASRadio.Checked -eq $true){SendAS}}
$AddSendOnRadio = GenRadioBox -x 185 -y 45 -width 130 -height 15 -text 'Send On Behalf + FA'
 $AddSendOnRadio.TabIndex = "4"
 $AddSendOnRadio.Add_CheckedChanged{
 if($AddSendOnRadio.Checked -eq $true){SendOn}}
$AddcalendarRadio = GenRadioBox -x 315 -y 45 -width 70 -height 15 -text 'Calendar'
 $AddcalendarRadio.TabIndex = "5"
 $AddcalendarRadio.Add_CheckedChanged{
 if($AddcalendarRadio.Checked -eq $true){Calendarfull}}

#$AddSearchBtn = GenerateButton -text 'Search' -x 300 -y 19 -action{SearchMember}
$AddGropBox = GenGroupBox -x 1 -y 1 -width 385 -height 70 
$Addlodiangsendas = GenerateLabel -text 'please wait...' -x 120 -y 150
$Addlodiangsendas.Font= New-Object System.Drawing.Font("ariel",22,[System.Drawing.FontStyle]::Regular)

$AddUserNameLABEL = GenerateLabel -text 'AddUserName:' -x 5 -y 87
$AddUserNameLABEL.Font= New-Object System.Drawing.Font("ariel",10,[System.Drawing.FontStyle]::Regular)
$AddUserNameINPUT = GenTextBox -x 107 -y 87 -width 180
$AddUserNameINPUT.Add_Click({
$AddUserNameINPUT.TextAlign = "left"
$AddUserNameINPUT.Text = ""
$AddUserNameINPUT.BackColor = "window"})

$AddUserNameINPUT.add_GotFocus({
$AddUserNameINPUT.TextAlign = "left"
$AddUserNameINPUT.Text = ""
$AddUserNameINPUT.BackColor = "window"})
$AddUserNameBTN = GenerateButton -text 'Add' -x 305 -y 85 -action {addUsrMail}
$AddGropBox2 = GenGroupBox -x 1 -y 70 -width 385 -height 50

$AddListmemberLABEL = GenerateLabel -text 'ListMemebers:' -x 30 -y 130
$AddListMemebers =  GenerateListBox -x 7 -y 145 -width 120 -height 214
$AddListMemebers.RightToLeft ="yes"
#$AddListMemebers.MultiColumn = "true"
$AddListMemebers.SelectionMode = "MultiSimple"

$AddListMemebersImportLABEL = GenerateLabel -text 'ImportMemebers:' -x 260 -y 130
$AddListMemebersImport = GenerateListBox -x 245 -y 145 -width 130 -height 100
$AddListMemebersImport.RightToLeft ="yes"
$AddListMemebersImport.SelectionMode = "MultiSimple"

$AddListRemoveLABEL = GenerateLabel -text 'RemoveMemebers:' -x 260 -y 245
$AddListRemove = GenerateListBox -x 245 -y 260 -width 130 -height 100
$AddListRemove.RightToLeft ="yes"

$AddSelAllBTN = GenerateButton -text 'Move All' -x 150 -y 160 -action {MoveAll}
$AddSelOneBTN = GenerateButton -text 'Move' -x 150 -y 200 -action {MoveAccount}
$AddImportBtN = GenerateButton -text 'Import'-x 150 -y 240 -action {Importfile}
$AddRemoveBTN = GenerateButton -text 'Remove' -x 150 -y 280 -action {RemoveMembers}
$AddExprotBTN = GenerateButton -text 'Export List' -x 150 -y 320 -action {ExportList}

$AddGropBox4 = GenGroupBox -x 1 -y 120 -width 385 -height 267
$AddcancelBTN = GenerateButton -text 'Exit' -x 60 -y 390 -action {if([System.Windows.Forms.MessageBox]::Show("Do you wish to close the program?",` 
        "Shared MailBox",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "no"){
        return
    }
   $FormMain.Close()}
$AddSaveBTN   = GenerateButton -text 'Save' -x 150 -y 390 -action {SaveAll}

$AddClearBTN = GenerateButton -text 'Clear' -x 240 -y 390 -action {ClearForm}
#--------------------------------------Accept Messge TAB--------------------------------------#
$MsgGropBox = GenGroupBox -x 1 -y 1 -width 385 -height 55 
$MsgAliasLABEL = GenerateLabel -text 'Alias:' -x 5 -y 20
 $MsgAliasLABEL.Font= New-Object System.Drawing.Font("ariel",13,[System.Drawing.FontStyle]::Regular)
$MsgAliasINPUT = GenTextBox -x 55 -y 22 -width 245
 $MsgAliasINPUT.text = "Alias Name Only" 
 $MsgAliasINPUT.TextAlign ="center"
 $MsgAliasINPUT.Add_Click({$MsgAliasINPUT.TextAlign = "left";$MsgAliasINPUT.Text = "";$MsgAliasINPUT.BackColor = "window"})
 $MsgAliasINPUT.add_GotFocus({$MsgAliasINPUT.TextAlign = "left";$MsgAliasINPUT.Text = "";$MsgAliasINPUT.BackColor = "window"})
$MsgAliasBTN=GenerateButton -x 305 -y 20 -text 'Search' -action{CheckAccpet}
 $MsgAliasBTN.AutoSize=$True

$MsgGropBox2 = GenGroupBox -x 1 -y 55 -width 385 -height 50
$MsgUserNameLABEL = GenerateLabel -text 'Add UserName:' -x 5 -y 70
 $MsgUserNameLABEL.Font= New-Object System.Drawing.Font("ariel",11,[System.Drawing.FontStyle]::Regular)
$MsgUserNameINPUT = GenTextBox -x 119 -y 70 -width 180
 $MsgUserNameINPUT.Add_Click({$MsgUserNameINPUT.TextAlign = "left";$MsgUserNameINPUT.Text = "";$MsgUserNameINPUT.BackColor = "window"})
 $MsgUserNameINPUT.add_GotFocus({$MsgUserNameINPUT.TextAlign = "left";$MsgUserNameINPUT.Text = "";$MsgUserNameINPUT.BackColor = "window"})
$MsgUserNameBTN = GenerateButton -text 'Add' -x 305 -y 70 -action{addUsrMail}

$MsgGropBox3 = GenGroupBox -x 1 -y 105 -width 385 -height 267
$MsgListmemberLABEL = GenerateLabel -text 'ListMemebers:' -x 15 -y 120
 $MsgListmemberLABEL.Font= New-Object System.Drawing.Font("ariel",11,[System.Drawing.FontStyle]::Regular)
$MsgListMemebers =  GenerateListBox -x 7 -y 145 -width 120 -height 214
 $MsgListMemebers.RightToLeft ="yes"
 $MsgListMemebers.SelectionMode = "MultiSimple"

$MsgSelAllBTN = GenerateButton -text 'Move All' -x 150 -y 160 -action {MoveAll}
$MsgSelOneBTN = GenerateButton -text 'Move' -x 150 -y 200 -action {MoveAccount}
$MsgImportBtN = GenerateButton -text 'Import'-x 150 -y 240 -action {Importfile}
$MsgRemoveBTN = GenerateButton -text 'Remove' -x 150 -y 280 -action {RemoveMembers}
$MsgExprotBTN = GenerateButton -text 'Export List' -x 150 -y 320 -action {ExportList}

$MsgListMemebersImportLABEL = GenerateLabel -text 'ImportMemebers:' -x 253 -y 120
 $MsgListMemebersImportLABEL.Font= New-Object System.Drawing.Font("ariel",11,[System.Drawing.FontStyle]::Regular)
$MsgListMemebersImport = GenerateListBox -x 245 -y 145 -width 130 -height 90
 $MsgListMemebersImport.RightToLeft ="yes"
 $MsgListMemebersImport.SelectionMode = "MultiSimple"

$MsgListRemoveLABEL = GenerateLabel -text 'RemoveMemebers:' -x 245 -y 235
 $MsgListRemoveLABEL.font= New-Object System.Drawing.Font("ariel",11,[System.Drawing.FontStyle]::Regular)
$MsgListRemove = GenerateListBox -x 245 -y 260 -width 130 -height 100
 $MsgListRemove.RightToLeft ="yes"


$MsgSaveBTN   = GenerateButton -text 'Save' -x 150 -y 390 -action {SaveAll}
$MsgClearBTN = GenerateButton -text 'Clear' -x 240 -y 390 -action {ClearAccpetTab}
$MsgExitBTN = GenerateButton -text 'Exit' -x 60 -y 390 -action {if([System.Windows.Forms.MessageBox]::Show("Do you wish to close the program?",` 
                                                                  "Shared MailBox",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "no"){return}
                                                                   $FormMain.Close()
                                                               }

#--------------------------------------Object Main form--------------------------------------#
$pat = "^[0-9]"
$hebrewname = "^[א-ת]"
$EnglishName= "^[a-z,A-Z]"
$AllPat = "^[0-9,א-ת,a-z,A-Z]"

$MsgTbObject=@($MsgUserNameLABEL,$MsgUserNameLABEL,$MsgUserNameINPUT,$MsgUserNameBTN,`
$MsgGropBox2,$MsgListmemberLABEL,$MsgListMemebers,$MsgSelAllBTN,$MsgSelOneBTN,$MSgImportBtN,$MsgRemoveBTN,$MsgExprotBTN,`
$MsgListMemebersImportLABEL,$MsgListMemebersImport,$MsgListRemoveLABEL,$MsgListRemove,$MsgGropBox3,$MsgSaveBTN)

$MsgTabOB=@($MsgAliasLABEL,$MsgAliasINPUT,$MsgAliasBTN,$MsgGropBox,$MsgExitBTN,$MsgClearBTN)
foreach($ctl in $MsgTabOB){
     $AcceptMsgTAB.Controls.Add($ctl)}

$NewTABControl = @($NewNameLABEL,$NewNameINPUT,$NewAliasLABEL,$NewAliasINPUT,$NewDbLABEL,$NewDbcombobox,$NewNamecheck,$NewAliascheck,$NewGropBox,$newWhomanageLABEL,$newWhomanageINPUT,$NewGropBox2,$NewclearBtn`
,$NewCencal,$NewCreateBtn);
  foreach($ctl in $NewTABControl){
     $NewTAB.Controls.Add($ctl)}
$AddTABshowControl=@($AddUserNameLABEL,$AddUserNameINPUT,$AddUserNameBTN,$AddGropBox2,`
$AddListMemebers,$AddListmemberLABEL,$AddListMemebersImport,$AddListMemebersImportLABEL,$AddRemoveBTN,$AddImportBtN,$AddSelOneBTN,$AddSelAllBTN,$AddExprotBTN,$AddListRemove,`
$AddListRemoveLABEL,$AddSaveBTN,$AddGropBox4)
$AddcalanderControl=@($AddUserNameLABEL,$AddUserNameINPUT,$AddUserNameBTN,$AddGropBox2,$AddSaveBTN,$AddListRemoveCalendar,`
$AddListRemoveCalendarLABEL,$AddListMemebersCalendarImport,$AddListMemebersCalendarImportLABEL,$AddListMemebersCalendar,`
$AddListMemebersCalendar,$AddListmemberCalendarLABEL,$AddremoveCalendar,$AddGropBox4)
$AddTABControl = @($AddAliasLABEL,$AddAliasINPUT,$AddSendASRadio,$AddSendOnRadio,$AddFullRadio,$AddcalendarRadio,$AddGropBox,$AddcancelBTN,$AddClearBTN)
  foreach($ctl in $AddTABControl){
     $AddTAB.Controls.Add($ctl)}

$ForMainControls = @($Tabctrl);
  foreach($ctl in $ForMainControls){
     $FormMain.Controls.Add($ctl)
}

$fromheb = @($hebcancelbtn,$hebChooseUserNameLABEL,$heblist,$hebchoosebtn);
  foreach($ctl in $fromheb){
     $Formherew.Controls.Add($ctl)
}

$fromcalendar = @($choosepermissionsLABEL,$choosepermissionslist,$choosepermissionscentelbtn,$choosepermissionsbtn);
  foreach($ctl in $fromcalendar){
     $CalendarMain.Controls.Add($ctl)
}

$fromalias = @($aliascancelbtn,$choosenameLABEL,$aliasLST,$aliaschoosebtn);
  foreach($ctl in $fromalias){
     $Formaliashebrew.Controls.Add($ctl)
}

$FormMain.ShowDialog()
