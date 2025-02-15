########################################################################################################
#                           Example of use of the EKARA API                                            #
########################################################################################################
# Swagger interface : https://api.ekara.ip-label.net                                                   #
# To be personalized in the code before use: username / password / API-KEY / Partners_only             #
# Purpose of the script : Collect monthly and annual EKARA consumptions then generate an Excel report  #
# Set-ExecutionPolicy Unrestricted                                                                     #
########################################################################################################
# Author : Guy Sacilotto
# Last Update : 15/02/2025
# Version : 1.1

<#
Authentication :  user / password / API-KEY
Method call : 
    auth/login  
    adm-api/clients
    adm-api/consumption

Restitution: Console / Excel
#>

Clear-Host

#region VARIABLES
#========================== SETTING THE VARIABLES ===============================
$error.clear()
add-type -AssemblyName "Microsoft.VisualBasic"
Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$global:API = "https://api.ekara.ip-label.net"                                                      # Webservice URL
$global:UserName = "xxxxxxxxxxxxxxxx"                                                               # EKARA Account
$global:PlainPassword = "xxxxxxxxxxxxxxxx"                                                          # EKARA Password
$global:API_KEY = ""                                                                                # EKARA Key account
$global:Partners_only = $false                                                                       # True / False
$global:ExcelVisible = $False

$global:Result_OK = 0
$global:Result_KO = 0
[String]$global:date = [DateTime]::Now.ToString("yyyy-MM-dd HH-mm-ss")                              # Recupere la date du jour

$global:headers = $null
$global:headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"             # Create Header
$headers.Add("Accept","application/json")                                                           # Setting Header
$headers.Add("Content-Type","application/json")                                                     # Setting Header
$global:clientId = ""

# Recherche le chemin du script
if ($psISE) {
    [String]$global:Path = Split-Path -Parent $psISE.CurrentFile.FullPath
    if($Debug -eq $true){Write-Host "Path ISE = $Path" -ForegroundColor Yellow}
} else {
    #[String]$global:Path = split-path -parent $MyInvocation.MyCommand.Path
    [String]$global:Path = (Get-Item -Path ".\").FullName
    if($Debug -eq $true){Write-Host "Path Direct = $Path" -ForegroundColor Yellow}
}

# Authentication choice
    # 1 = Without asking for an account and password (you must configure the account and password in this script.)
    # 2 = Request the entry of an account and a password (default)
    # 3 = With API-KEY
    $global:Auth = 2
#endregion

#region Functions
function Authentication{
    try{
        Switch($Auth){
            1{
                # Without asking for an account and password
                if(($null -ne $UserName -and $null -ne $PlainPassword) -and ($UserName -ne '' -and $PlainPassword -ne '')){
                    Write-Host "--- Automatic AUTHENTICATION (account) ---------------------------" -BackgroundColor Green
                    $uri = "$API/auth/login"                                                                                                    # Webservice Methode
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Verbose -Body @{ email = "$UserName"; password = "$PlainPassword"}    # Call WebService method
                    $global:Token = $response.token                                                                                             # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")                                                                        # Adding the TOKEN into header
                }Else{
                    Write-Host "--- Account and Password not set ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure the account and password in this script." -ForegroundColor Red
                    Break Script
                }
            }
            2{
                # Requests the entry of an account and a password (default) 
                Write-Host "------------------------------ AUTHENTICATION with account entry ---------------------------" -ForegroundColor Green
                $MyAccount = $Null
                $MyAccount = Get-credential -Message "EKARA login account" -ErrorAction Stop                                            # Request entry of the EKARA Account
                if(($null -ne $MyAccount) -and ($MyAccount.password.Length -gt 0)){
                    $UserName = $MyAccount.GetNetworkCredential().username
                    $PlainPassword = $MyAccount.GetNetworkCredential().Password
                    $uri = "$API/auth/login"
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Body @{ email = "$UserName"; password = "$PlainPassword"} -Verbose     # Call WebService method
                    $Token = $response.token                                                                                            # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")
                }Else{
                    Write-Host "--- Account and password not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must enter Account and password." -ForegroundColor Red
                    Break Script
                }
            }
            3{
                # With API-KEY
                Write-Host "------------------------------ AUTHENTICATION With API-KEY ---------------------------" -ForegroundColor Green
                if(($null -ne $API_KEY) -and ($API_KEY -ne '')){
                    $global:headers.Add("X-API-KEY", $API_KEY)
                }Else{
                    Write-Host "--- API-KEY not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure API-KEY." -ForegroundColor Red
                    Break Script
                }
            }
        }
        Write-Host "-------------------------------------------------------------" -ForegroundColor green
    }Catch{

    Write-Host "-------------------------------------------------------------" -ForegroundColor red 
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Break Script
    }
}

function Hide-Console{
    # .Net methods Permet de réduire la console PS dans la barre des tâches
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 2) | Out-Null                     #0 hide / 1 normal / 2 réduit 
}

Function List_Clients{
    #========================== List all customers =============================
    try{
        Write-Host "------------------- List all customers  ---------------------" -BackgroundColor Green
        $uri ="$API/adm-api/clients"
        $clients = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers                # All accounts

        if($Partners_only -eq $true){
            $clients = $clients | Where-Object{$_.isPartner -eq $True}                       # All partners
            $count = $clients.count
            Write-Host ("--> ["+$count+"] Partners ---------------------------") -ForegroundColor Blue
        }else{
            $count = $clients.count
            Write-Host ("--> ["+$count+"] customers ---------------------------") -ForegroundColor Blue
        }
     
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        function ListIndexChanged { 
            #$label2.Text = $listbox.SelectedItems.Count
            $okButton.enabled = $True
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'List all customers'
        $form.Size = New-Object System.Drawing.Size(350,500)
        $form.StartPosition = 'CenterScreen'
        $Form.Opacity = 1.0
        $Form.TopMost = $false
        $Form.ShowIcon = $true                                              # Enable icon (upper left corner) $ true, disable icon
        #$Form.FormBorderStyle = 'Fixed3D'                                  # bloc resizing form
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(80,430)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = 'OK'
        $okButton.AutoSize = $true
        $okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom 
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $okButton.enabled = $False
        $form.AcceptButton = $okButton

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(180,430)
        $cancelButton.Size = New-Object System.Drawing.Size(75,23)
        $cancelButton.Text = 'Cancel'
        $cancelButton.AutoSize = $true
        $cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom 
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $cancelButton
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = 'Select the customers to run Quota inventory:'
        $label.AutoSize = $true
        $label.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
        -bor [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left `
        -bor [System.Windows.Forms.AnchorStyles]::Right

        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(10,435)
        $label2.Size = New-Object System.Drawing.Size(20,20)
        $label2.Text = $count
        $label2.AutoSize = $true
        $label2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left 

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10,40)
        $listBox.Size = New-Object System.Drawing.Size(310,20)
        $listBox.Height = 380
        $listBox.SelectionMode = 'One'
        $ListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
        -bor [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left `
        -bor [System.Windows.Forms.AnchorStyles]::Right

        $listboxCollection =@()

        foreach($client in $clients){
            $Object = New-Object Object 
            $Object | Add-Member -type NoteProperty -Name id -Value $client.id
            $Object | Add-Member -type NoteProperty -Name name -Value $client.name
            $listboxCollection += $Object
        }
        
        # Count selected item
        $ListBox.Add_SelectedIndexChanged({ ListIndexChanged })

        #Add collection to the $listbox
        $listBox.Items.AddRange($listboxCollection)
        $listBox.ValueMember = "$listboxCollection.id"
        $listBox.DisplayMember = "$listboxCollection.name"
        
        #Add composant into Form
        $form.Controls.Add($okButton)
        $form.Controls.Add($cancelButton)
        $form.Controls.Add($listBox)
        $form.Controls.Add($label2)
        $form.Controls.Add($label)
        $form.Topmost = $true
        $result = $form.ShowDialog()
        
        if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and $listbox.SelectedItems.Count -gt 0)
        {
            $ItemsName = $listBox.SelectedItems.name
            $global:ItemsID = $listBox.SelectedItems.id
            $global:clientId = $ItemsID
            Write-Host "--> Client name selected :$ItemsName (ID = $clientId)" -ForegroundColor Blue

        }else{
            Write-Host "No customer selected !" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(`
                "------------------------------------`n`r No customer selected !`n`r------------------------------------`n`r",`
                "Resultat",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Warning)
            Break Script
        }
    }
    catch{
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Break Script
    }
}

function List_subAccount($clientId){
    #========================== List account ans sub account =============================
    try{
        # Creation du ficheir XLS ---------------------------------------------------
        Write-Host "------------------------------------------------------------------------" -ForegroundColor White
        Write-Host ("--> Create XLS file") -ForegroundColor Green
        [String]$global:NewFile = 'EKARA_CONSO_'+$clientName+"_"+$date + '.xlsx'                #  Nom du Nouveau fichier XLS 
        Write-Host "--> File name ["$NewFile"]" -ForegroundColor Blue  
        Write-Host "--> Folder ["$path"]" -ForegroundColor Blue         
        $ExcelPath = "$Path\$NewFile"                                                           # Nom et Chemin du fichier
        $objExcel = New-Object -ComObject Excel.Application                                     # Creation du fichier XLS
        $objExcel.Visible = $ExcelVisible                                                       # Affiche Excel

        $FirstRowColor = 20                                                                     # Cellule Couleur 1
        $SecondtRowColor = 2                                                                    # Cellule Couleur 2
        $color = $FirstRowColor
        $VAlignTop = -4160                                                                      # Cellule Vertical Aligne Haut
        $VAlignBottom = -4107                                                                   # Cellule Vertical Aligne bas
        $VAlignCenter = -4108                                                                   # Cellule Vertical Aligne Center
        $HAlignCenter = -4108                                                                   # Cellule Horizontal Aligne Center
        $HAlignLeft = -4131                                                                     # Cellule Horizontal Aligne Gauche
        $HAlignRight = -4152                                                                    # Cellule Horizontal Aligne Droite

        # Microsoft.Office.Interop.Excel.XlBorderWeight
        $xlHairline = 1
        $xlThin = 2
        $xlThick = 4
        $xlMedium = -4138

        # Microsoft.Office.Interop.Excel.XlBordersIndex
        $xlDiagonalDown = 5
        $xlDiagonalUp = 6
        $xlEdgeLeft = 7
        $xlEdgeTop = 8
        $xlEdgeBottom = 9
        $xlEdgeRight = 10
        $xlInsideVertical = 11
        $xlInsideHorizontal = 12

        # Microsoft.Office.Interop.Excel.XlLineStyle
        $xlContinuous = 1
        $xlDashDot = 4
        $xlDashDotDot = 5
        $xlSlantDashDot = 13
        $xlLineStyleNone = -4142
        $xlDouble = -4119
        $xlDot = -4118
        $xlDash = -4115

        $xlAutomatic = -4105
        $xlBottom = -4107
        $xlCenter = -4108
        $xlContext = -5002
        $xlNone = -4142

        # color index
        $xlColorIndexBlue = 23                                                                  # <-- depends on default palette

        $row = 1                                                                                # Set first row
        $Column = 1                                                                             # Set first coumn
        $workbook = $objExcel.Workbooks.add()                                                   # Ajout d'une feuille au fichier XLS
        $CurrentSheet = $workbook.WorkSheets.item(1)                                            # Selection du nouvel onglet                    
        $CurrentSheet.Name = "CONSO"                                                            # Nom de l'onglet

        #========================== adm-api/clients =============================
        Write-Host "------------------- List all sub Account of customer [$clientId] ---------------------" -BackgroundColor Green
        $uri =($API+"/adm-api/clients?clientId="+$clientId)
        $client = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers                    # Account and Sub account
        $client = $client | Sort-Object -Property isPartner, name -Descending                   # Order with Partner in first
        $NB_Client = $client.count
        
        $global:Table = Table                                                                   # Create table with letters and numbers
        
        Write-Host "Number of accounts and sub-accounts [$NB_Client]"
        
        foreach ($account in $client){
            $account_Name = $account.name                                                                       # Remember account name
            $account_Quota = $account.quota                                                                     # Remember account Quota
            if($account.isActif -eq 0){$accountStatus="Inactif"}else{$accountStatus="Actif"}                    # Remember account status
            if($account.isPartner -eq $True){$accountType="Partenaire"}else{$accountType="Sous-compte"}         # Remember account type
            Write-Host ("--> Account Name [$account_Name] status : [$accountStatus] Type : [$accountType]") -ForegroundColor Blue

            $Monthly_equivalence = [int][Math]::Round($account_Quota*24*30.42,[MidpointRounding]::AwayFromZero)
            $Annual_equivalence = [int][Math]::Round($account_Quota*24*30.42*12,[MidpointRounding]::AwayFromZero)
            Write-Host ("--> Account Quota [$account_Quota] / Monthly equivalence ["+ ($Monthly_equivalence) +"] / Annual equivalence ["+ ($Annual_equivalence) +"]" ) -ForegroundColor Blue

           <#
            $uri =($API+"/adm-api/tag/values?")
            $detailClient = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ContentType 'application/json' -Body @{ type = "client"; id = "$account"; clientId = "$clientId"}                    # Account ID and Partner ID
            $Sector = $detailClient[7].value
            $TypeService = $detailClient[8].value
            Write-Host("Sector : [" + $Sector + "] / TypeService : [" +$TypeService +"]")
            #>

            # Search conso
            $conso = get_conso -clientId $account.id -clientName $account_Name
            
            #region XLS
            #------------------------------------------------------------------------------------------------------------------------------------
            # Creation du titre (premiere ligne)
            $color = $FirstRowColor
            $Column++
            $CurrentSheet.Cells.Item($row,$Column) = "Synthese du compte"
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+11)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))

            $range.Select() | Out-Null
            $range.Font.Size = 18
            $range.Font.Bold = $True
            $range.Font.ColorIndex = 2
            $range.MergeCells = $true                                                                           # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            $range.BorderAround($xlContinuous,$xlThick,$xlColorIndexBlue) | Out-Null
            $range.Interior.ColorIndex = 14                                                                     # Change la couleur des cellule
            
            # Creation du contenu
            $row++
            $CurrentSheet.Cells.Item($row,$Column) = "Name"
            $CurrentSheet.Cells.Item($row,$Column).Font.Bold = $True
            
            $Column++
            $CurrentSheet.Cells.Item($row,$Column) = $account_Name
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+10)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))
            $range.Select() | Out-Null
            $range.MergeCells = $true                                                                           # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            $Column--
            
            $row++
            $CurrentSheet.Cells.Item($row,$Column) = "Status"
            $CurrentSheet.Cells.Item($row,$Column).Font.Bold = $True

            $Column++
            if ($accountStatus -eq "Inactif" ){
                $CurrentSheet.Cells.Item($row,$Column).Font.ColorIndex = 3                          # Set font color (3 = red)
                $CurrentSheet.Cells.Item($row,$Column) = [string]$accountStatus                     # Add data
                
            }else{
                #$finalWorkSheet.Cells.Item($row,$Column).Font.ColorIndex = 1                       # Set font color (1 = black)
                $CurrentSheet.Cells.Item($row,$Column) = [string]$accountStatus                     # Add data
            }
            
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+10)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))
            $range.Select() | Out-Null
            $range.MergeCells = $true                                                                   # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            $Column--

            $row++
            $CurrentSheet.Cells.Item($row,$Column) = "Type"
            $CurrentSheet.Cells.Item($row,$Column).Font.Bold = $True

            $Column++
            $CurrentSheet.Cells.Item($row,$Column) = $accountType
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+10)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))
            $range.Select() | Out-Null
            $range.MergeCells = $true                                                                   # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            $Column--

            $row++
            $CurrentSheet.Cells.Item($row,$Column) = "Quota"
            $CurrentSheet.Cells.Item($row,$Column).Font.Bold = $True

            $Column++
            $CurrentSheet.Cells.Item($row,$Column) = $account_Quota
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+10)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))
            $range.Select() | Out-Null
            $range.MergeCells = $true                                                                                   # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            if([int]$account_Quota.count -gt 0){  
                $Monthly_equivalence = [int][Math]::Round($account_Quota*24*30.42,[MidpointRounding]::AwayFromZero)
                $Annual_equivalence = [int][Math]::Round($account_Quota*24*30.42*12,[MidpointRounding]::AwayFromZero)
                [String]$equivalence = ("Equivalence du quota`n - Mensuelle = "+ $Monthly_equivalence + "`n - Annuelle = "+$Annual_equivalence)
                [void]$CurrentSheet.Cells.Item($row,$Column).AddComment(""+$equivalence+"")                             # Add data into comment
                $CurrentSheet.Cells.Item($row,$Column).Comment.Shape.TextFrame.Characters().Font.Size = 8               # Format comment
                $CurrentSheet.Cells.Item($row,$Column).Comment.Shape.TextFrame.Characters().Font.Bold = $False          # Format comment                                                                                        
            }
            $Column--

            $row = $row + 2
            $CurrentSheet.Cells.Item($row,$Column) = "Execution mensuelle"
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+11)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))

            $range.Select() | Out-Null
            $range.Font.Size = 18
            $range.Font.Bold = $True
            $range.Font.ColorIndex = 2
            $range.MergeCells = $true                                                                           # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            $range.BorderAround($xlContinuous,$xlThick,$xlColorIndexBlue) | Out-Null
            $range.Interior.ColorIndex = 14                                                                     # Change la couleur des cellule

            $row++

            for($i=0 ; $i -lt $conso.length ; $i++){
                $CurrentSheet.Cells.Item($row,$Column) = $conso[$i].Month
                $CurrentSheet.Cells.Item($row,$Column).Font.Bold = $True
                $Column++
            }

            $Column = $Column-$conso.length
            $row++
            for($i=0 ; $i -lt $conso.length ; $i++){
                $CurrentSheet.Cells.Item($row,$Column) = $conso[$i].Conso
                $Column++
            }

            $Column = $Column-$conso.length
            $row = $row + 2

            $CurrentSheet.Cells.Item($row,$Column) = "Execution Annuelle"
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+11)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))

            [string]$rangeChart = ($letterStart.Letters+($row-3))                                                   # Remember position for the chart range

            $range.Select() | Out-Null
            $range.Font.Size = 18
            $range.Font.Bold = $True
            $range.Font.ColorIndex = 2
            $range.MergeCells = $true                                                                           # Fusionne les cellules
            $range.HorizontalAlignment = -4108
            $range.BorderAround($xlContinuous,$xlThick,$xlColorIndexBlue) | Out-Null
            $range.Interior.ColorIndex = 14                                                                     # Change la couleur des cellule

            $row++
            $CurrentSheet.Cells.Item($row,$Column) = ($conso | Measure-Object Conso -Sum).Sum
            $letterStart = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq $Column} | select-object -Property Letters
            $letterEnd = $Table | Select-Object -Property Letters, Num | Where-Object {$_.Num -eq ($Column+11)} | select-object -Property Letters
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+$row))
            $range.Select() | Out-Null
            $range.MergeCells = $true                                                                           # Fusionne les cellules
            $range.HorizontalAlignment = -4108

            $row++
            $row++

            $DataforChart = $CurrentSheet.Range($rangeChart).CurrentRegion                              # Set range Data for the chart

            # Add Chart 
            $xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
            $chart = $CurrentSheet.Shapes.AddChart().Chart                                              # Adding the Charts
            $chart.chartType=$xlChart::xlLineMarkersStacked                                             # Providing the chart types
            $chart.SetSourceData($DataforChart)                                                         # Providing the source data
            
            $chart.HasTitle = $true                                                                     # modify the chart title
            $chart.ChartTitle.Text = "Consomation"                                                      # modify the chart title
            $chart.HasLegend = $False                                                                   # Delete the chart Legend

            # Setting up the position of chart
            $range = $CurrentSheet.Range(($letterStart.Letters+$row),($letterEnd.Letters+($row+12)))    # Selecte range for position of chart
            $ChtOb = $chart.Parent                                                                      # Selecte chart
            $ChtOb.Top = $range.Top                                                                     # Set Top position of chart
            $ChtOb.Left = $range.Left                                                                   # Set Left position of chart
            $ChtOb.Height = $range.Height                                                               # Set Height position of chart
            $ChtOb.Width = $range.Width                                                                 # Set Width position of chart

            $row = 1
            $Column = $Column+$conso.length
            #------------------------------------------------------------------------------------------------------------------------------------
            #endregion
        }
        
        $workbook.Worksheets.Item(1).UsedRange.Cells.EntireColumn.AutoFit() | Out-Null 
        
        # Enregistrement du fichier ------------------------------------------------------------------------
        Write-Host ("--> Saving data Excel file. [" + $Path + "\" + $NewFile +  "]") -ForegroundColor Blue
        $workbook.SaveAs($ExcelPath) | Out-Null                                                                 # Save file
        # Fermeture du fichier 
        Write-Host ("--> Closing Excel file. [" + $Path + "\" + $NewFile +  "]") -ForegroundColor Blue
        $workbook.Close() | Out-Null                                                                            # Close Workbook
        $objExcel.quit() | Out-Null                                                                             # Close Excel
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

        # Ouvre le fichier XLS
        Write-Host "--> Open Excel file." -ForegroundColor Blue
        Invoke-Item -Path "$Path\$NewFile"

        Write-Host ("END of Partner Quota Inventory ["+$clientId+"]: " + $Result_OK) -BackgroundColor Green
        [System.Windows.Forms.MessageBox]::Show("Quota inventory [$clientId] : $Result_OK","Resultat",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch{
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Break Script
    }
}

function get_conso($clientId, $clientName){
    #========================== Consumption inventory =============================
    try{
        write-Host "--- Get CONSO for id client [$clientId] [$clientName]---------------------------" -BackgroundColor Green
        $uri ="$API/adm-api/consumption?clientId=$clientId"
        $conso = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -Verbose
        Write-Host ("Nb Plugin : " + ($conso.consumption).count) -ForegroundColor Cyan
        $consumption = @()
        #Start-Sleep -Seconds 1.5
        # For each plugin
        for($i=0 ; $i -lt $conso.consumption.length ; $i++){
            #For each consumption
            for($j=0 ; $j -lt $conso.consumption[$i].data.x.length ; $j++){
                $Object = New-Object Object 
                $Object | Add-Member -type NoteProperty -Name Plugin -Value $conso.consumption[$i].name
                $Object | Add-Member -type NoteProperty -Name Month -Value (Get-Date $conso.consumption[$i].data.x[$j] -uformat "%b")
                $Object | Add-Member -type NoteProperty -Name Conso -Value $conso.consumption[$i].data.y[$j]
                $consumption += $Object
                Write-Host "." -NoNewline -ForegroundColor Blue
            }
        }

        $global:listeConso = @()
        $consumption | Group-Object Month | ForEach-Object {
            $Object = New-Object psobject -Property @{
                Month = $_.Name
                Conso = ($_.Group  | Measure-Object Conso -Sum).Sum
            }
            $listeConso += $Object
        }
        
        $yearConso = ($listeConso | Measure-Object Conso -Sum).Sum                                                  # Calculate annual consumption
        $listeConso = $listeConso | Sort-Object -Property @{Expression={(++$script:i)}; Descending=$true}           # Reverse array
        #$listeConso | Out-GridView -Title "CONSO for [$clientName]"                                                 # Show results
        
        Write-Host ""
        Write-Host "---------------------------------------" -ForegroundColor Blue
        for($i=0 ; $i -lt $listeConso.length ; $i++){
            Write-host ("--> Month : [" + $listeConso[$i].Month + "] Conso : [" + $listeConso[$i].Conso + "]") -ForegroundColor Blue
        }

        Write-Host "---------------------------------------" -ForegroundColor Blue
        Write-host "--> Annual consumption : $yearConso" -ForegroundColor Blue
        Write-Host ""
        Write-Host "-------------------------------------------------------------" -ForegroundColor green

        return $listeConso
    }
    catch{
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
    }
}

function Table{
    $Table = @()
    for ($test0 = 0; $test0 -lt 26; $test0++){
        $letter = [char](65 + $test0)
        $num = $test0+1
        $Object = New-Object Object 
        $Object | Add-Member -type NoteProperty -Name Letters -Value $letter
        $Object | Add-Member -type NoteProperty -Name Num -Value $num
        $Table += $Object
    }
    for ($test1 = 0; $test1 -lt 26; $test1++){
        $letter = [char](65 + 0)+[char](65 + $test1)
        $num = $test0+$test1+1
        $Object = New-Object Object 
        $Object | Add-Member -type NoteProperty -Name Letters -Value $letter
        $Object | Add-Member -type NoteProperty -Name Num -Value $num
        $Table += $Object
    }
    for ($test2 = 0; $test2 -lt 26; $test2++){
        $letter = [char](65 + 1)+[char](65 + $test2)
        $num = $test0+$test1+$test2+1
        $Object = New-Object Object 
        $Object | Add-Member -type NoteProperty -Name Letters -Value $letter
        $Object | Add-Member -type NoteProperty -Name Num -Value $num
        $Table += $Object
    }
    for ($test3 = 0; $test3 -lt 26; $test3++){
        $letter = [char](65 + 2)+[char](65 + $test3)
        $num = $test0+$test1+$test2+$test3+1
        $Object = New-Object Object 
        $Object | Add-Member -type NoteProperty -Name Letters -Value $letter
        $Object | Add-Member -type NoteProperty -Name Num -Value $num
        $Table += $Object
    }
    for ($test4 = 0; $test4 -lt 26; $test4++){
        $letter = [char](65 + 3)+[char](65 + $test4)
        $num = $test0+$test1+$test2+$test3+$test4+1
        $Object = New-Object Object 
        $Object | Add-Member -type NoteProperty -Name Letters -Value $letter
        $Object | Add-Member -type NoteProperty -Name Num -Value $num
        $Table += $Object
    }
    for ($test5 = 0; $test5 -lt 26; $test5++){
        $letter = [char](65 + 4)+[char](65 + $test5)
        $num = $test0+$test1+$test2+$test3+$test4+$test5+1
        $Object = New-Object Object 
        $Object | Add-Member -type NoteProperty -Name Letters -Value $letter
        $Object | Add-Member -type NoteProperty -Name Num -Value $num
        $Table += $Object
    }
    return $Table
}
#endregion

#region Main
#========================== START SCRIPT ======================================
Authentication
List_Clients
List_subAccount -clientId $clientId
#endregion