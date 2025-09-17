# Modülleri yükle
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Form oluştur
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Sorgu Aracı"
$form.Size = New-Object System.Drawing.Size(1200,800) # Genişlik artırıldı
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None

# ToolTip oluştur
$toolTip = New-Object System.Windows.Forms.ToolTip

# GroupBox: Domain ve Credential
$groupDomain = New-Object System.Windows.Forms.GroupBox
$groupDomain.Location = New-Object System.Drawing.Point(10,10)
$groupDomain.Size = New-Object System.Drawing.Size(1160,80) # Genişlik artırıldı
$groupDomain.Text = "Domain ve Kimlik Bilgileri"
$form.Controls.Add($groupDomain)

# Domain girişi için Label ve TextBox
$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Location = New-Object System.Drawing.Point(10,30)
$lblDomain.Size = New-Object System.Drawing.Size(150,20)
$lblDomain.Text = "Domain Adı:"
$groupDomain.Controls.Add($lblDomain)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = New-Object System.Drawing.Point(170,30)
$txtDomain.Size = New-Object System.Drawing.Size(200,20)
$txtDomain.Text = (Get-ADDomain).DNSRoot
$toolTip.SetToolTip($txtDomain, "Active Directory domain adını girin")
$groupDomain.Controls.Add($txtDomain)

# Credential Butonu
$btnCredential = New-Object System.Windows.Forms.Button
$btnCredential.Location = New-Object System.Drawing.Point(380,25)
$btnCredential.Size = New-Object System.Drawing.Size(150,30)
$btnCredential.Text = "Credential Gir"
$toolTip.SetToolTip($btnCredential, "Farklı bir kullanıcıyla bağlanmak için kimlik bilgisi girin")
$btnCredential.Add_Click({
    try {
        $global:cred = Get-Credential -Message "Farklı domain için kullanıcı adı ve şifre girin."
        if ($global:cred) {
            [System.Windows.Forms.MessageBox]::Show("Credential başarıyla girildi.")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Credential girilmedi, mevcut kullanıcı kullanılacak.")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata: Credential girme başarısız! $_")
    }
})
$groupDomain.Controls.Add($btnCredential)

# GroupBox: OU ve Bilgisayar Seçimi
$groupOU = New-Object System.Windows.Forms.GroupBox
$groupOU.Location = New-Object System.Drawing.Point(10,100)
$groupOU.Size = New-Object System.Drawing.Size(1160,80)
$groupOU.Text = "OU ve Bilgisayar Seçimi"
$form.Controls.Add($groupOU)

# OU ComboBox ve Yükle Butonu
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Location = New-Object System.Drawing.Point(10,30)
$lblOU.Size = New-Object System.Drawing.Size(150,20)
$lblOU.Text = "OU Seç:"
$groupOU.Controls.Add($lblOU)

$comboOU = New-Object System.Windows.Forms.ComboBox
$comboOU.Location = New-Object System.Drawing.Point(170,30)
$comboOU.Size = New-Object System.Drawing.Size(200,20)
$comboOU.DropDownStyle = "DropDownList"
$toolTip.SetToolTip($comboOU, "Sorguları filtrelemek için bir OU seçin")
$groupOU.Controls.Add($comboOU)

$btnLoadOU = New-Object System.Windows.Forms.Button
$btnLoadOU.Location = New-Object System.Drawing.Point(380,25)
$btnLoadOU.Size = New-Object System.Drawing.Size(150,30)
$btnLoadOU.Text = "OU'ları Yükle"
$toolTip.SetToolTip($btnLoadOU, "Domaindeki OU'ları yüklemek için tıklayın")
$btnLoadOU.Add_Click({
    $domain = $txtDomain.Text
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen bir domain adı girin!")
        return
    }
    try {
        $params = @{ Filter = '*'; Server = $domain; Properties = 'DistinguishedName' }
        if ($global:cred) { $params['Credential'] = $global:cred }
        $ous = Get-ADOrganizationalUnit @params
        $comboOU.Items.Clear()
        $comboOU.Items.Add("Tüm Domain (OU Yok)")
        foreach ($ou in $ous) {
            $comboOU.Items.Add($ou.DistinguishedName)
        }
        [System.Windows.Forms.MessageBox]::Show("OU'lar yüklendi: " + $ous.Count + " adet.")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata: OU yükleme başarısız! $_")
    }
})
$groupOU.Controls.Add($btnLoadOU)

# Bilgisayar ComboBox ve Yükle Butonu
$lblComputers = New-Object System.Windows.Forms.Label
$lblComputers.Location = New-Object System.Drawing.Point(540,30)
$lblComputers.Size = New-Object System.Drawing.Size(150,20)
$lblComputers.Text = "Bilgisayar Seç:"
$groupOU.Controls.Add($lblComputers)

$comboComputers = New-Object System.Windows.Forms.ComboBox
$comboComputers.Location = New-Object System.Drawing.Point(700,30)
$comboComputers.Size = New-Object System.Drawing.Size(150,20)
$comboComputers.DropDownStyle = "DropDownList"
$toolTip.SetToolTip($comboComputers, "Kapatma veya NetBIOS işlemi için bir bilgisayar seçin")
$groupOU.Controls.Add($comboComputers)

$btnLoadComputers = New-Object System.Windows.Forms.Button
$btnLoadComputers.Location = New-Object System.Drawing.Point(860,25)
$btnLoadComputers.Size = New-Object System.Drawing.Size(150,30)
$btnLoadComputers.Text = "Bilgisayarları Yükle"
$toolTip.SetToolTip($btnLoadComputers, "Domaindeki bilgisayarları yüklemek için tıklayın")
$btnLoadComputers.Add_Click({
    $domain = $txtDomain.Text
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen bir domain adı girin!")
        return
    }
    try {
        $params = @{ 
            Filter = '*'
            Server = $domain
            Properties = 'Name'
            ResultPageSize = 1000
            ErrorAction = 'Stop'
        }
        if ($comboOU.SelectedItem -and $comboOU.SelectedItem -ne "Tüm Domain (OU Yok)") { 
            $params['SearchBase'] = $comboOU.SelectedItem 
        }
        if ($global:cred) { 
            $params['Credential'] = $global:cred 
        }
        Write-Host "Bilgisayarlar yükleniyor: Domain=$domain, SearchBase=$($params['SearchBase'])"
        $computers = Get-ADComputer @params | Sort-Object Name
        $comboComputers.Items.Clear()
        if ($computers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Seçilen OU veya domainde bilgisayar bulunamadı!")
            return
        }
        foreach ($computer in $computers) {
            $comboComputers.Items.Add($computer.Name)
        }
        [System.Windows.Forms.MessageBox]::Show("Bilgisayarlar yüklendi: " + $computers.Count + " adet.")
    } catch {
        Write-Host "Hata: $_"
        [System.Windows.Forms.MessageBox]::Show("Hata: Bilgisayar yükleme başarısız!`nDetay: $_")
    }
})
$groupOU.Controls.Add($btnLoadComputers)

# GroupBox: Sorgular
$groupQueries = New-Object System.Windows.Forms.GroupBox
$groupQueries.Location = New-Object System.Drawing.Point(10,190)
$groupQueries.Size = New-Object System.Drawing.Size(1160,150)
$groupQueries.Text = "Sorgular"
$form.Controls.Add($groupQueries)

# NumericUpDown: Şifre değiştirme sorgusu için gün sayısı
$lblPwdDays = New-Object System.Windows.Forms.Label
$lblPwdDays.Location = New-Object System.Drawing.Point(10,30)
$lblPwdDays.Size = New-Object System.Drawing.Size(150,20)
$lblPwdDays.Text = "Son Şifre Değiştirme (Gün):"
$groupQueries.Controls.Add($lblPwdDays)

$numPwdDays = New-Object System.Windows.Forms.NumericUpDown
$numPwdDays.Location = New-Object System.Drawing.Point(170,30)
$numPwdDays.Size = New-Object System.Drawing.Size(60,20)
$numPwdDays.Minimum = 1
$numPwdDays.Maximum = 365
$numPwdDays.Value = 30
$toolTip.SetToolTip($numPwdDays, "Şifre değiştirme süresini gün cinsinden belirtin")
$groupQueries.Controls.Add($numPwdDays)

# Buton: Son X günde şifre değiştirmemiş kullanıcılar
$btnPwd = New-Object System.Windows.Forms.Button
$btnPwd.Location = New-Object System.Drawing.Point(240,25)
$btnPwd.Size = New-Object System.Drawing.Size(150,30)
$btnPwd.Text = "Şifre Değiştirmemiş Kullanıcılar"
$toolTip.SetToolTip($btnPwd, "Belirtilen süre içinde şifre değiştirmemiş kullanıcıları listele")
$btnPwd.Add_Click({
    $domain = $txtDomain.Text
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen bir domain adı girin!")
        return
    }
    $searchBase = if ($comboOU.SelectedItem -and $comboOU.SelectedItem -ne "Tüm Domain (OU Yok)") { $comboOU.SelectedItem } else { $null }
    $params = @{ 
        Filter = {pwdLastSet -lt $date}
        Server = $domain
        Properties = @('Name', 'SamAccountName', 'pwdLastSet', 'Enabled')
        ErrorAction = 'Stop'
    }
    if ($searchBase) { $params['SearchBase'] = $searchBase }
    if ($global:cred) { $params['Credential'] = $global:cred }
    try {
        $days = [int]$numPwdDays.Value
        $date = (Get-Date).AddDays(-$days)
        $global:results = Get-ADUser @params | 
                          Select-Object Name, SamAccountName, @{Name='pwdLastSet';Expression={[datetime]::FromFileTime($_.pwdLastSet)}}, Enabled
        $dataGrid.DataSource = $global:results
        if ($global:results) {
            $htmlContent = $global:results | ConvertTo-Html -Title "AD Sorgu Sonuçları" -PreContent "<h2>Son $days Günde Şifre Değiştirmemiş Kullanıcılar</h2>" -Head $cssStyle
            $webBrowser.DocumentText = $htmlContent
        } else {
            $webBrowser.DocumentText = "<html><body><h2>Sonuç bulunamadı</h2></body></html>"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata: Sorgu başarısız! $_")
        $webBrowser.DocumentText = "<html><body><h2>Hata: Sorgu başarısız</h2></body></html>"
    }
})
$groupQueries.Controls.Add($btnPwd)

# NumericUpDown: AdminCount=1 sorgusu için (isteğe bağlı)
$lblAdmin1Days = New-Object System.Windows.Forms.Label
$lblAdmin1Days.Location = New-Object System.Drawing.Point(10,60)
$lblAdmin1Days.Size = New-Object System.Drawing.Size(150,20)
$lblAdmin1Days.Text = "AdminCount=1 için Gün:"
$groupQueries.Controls.Add($lblAdmin1Days)

$numAdmin1Days = New-Object System.Windows.Forms.NumericUpDown
$numAdmin1Days.Location = New-Object System.Drawing.Point(170,60)
$numAdmin1Days.Size = New-Object System.Drawing.Size(60,20)
$numAdmin1Days.Minimum = 0
$numAdmin1Days.Maximum = 365
$numAdmin1Days.Value = 0
$toolTip.SetToolTip($numAdmin1Days, "0 girerseniz gün filtresi uygulanmaz")
$groupQueries.Controls.Add($numAdmin1Days)

# Buton: AdminCount=1 olan kullanıcılar
$btnAdmin1 = New-Object System.Windows.Forms.Button
$btnAdmin1.Location = New-Object System.Drawing.Point(240,55)
$btnAdmin1.Size = New-Object System.Drawing.Size(150,30)
$btnAdmin1.Text = "AdminCount=1 Kullanıcılar"
$toolTip.SetToolTip($btnAdmin1, "AdminCount=1 olan kullanıcıları listele")
$btnAdmin1.Add_Click({
    $domain = $txtDomain.Text
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen bir domain adı girin!")
        return
    }
    $searchBase = if ($comboOU.SelectedItem -and $comboOU.SelectedItem -ne "Tüm Domain (OU Yok)") { $comboOU.SelectedItem } else { $null }
    $params = @{ 
        Server = $domain
        Properties = @('Name', 'SamAccountName', 'adminCount', 'Enabled')
        ErrorAction = 'Stop'
    }
    if ($searchBase) { $params['SearchBase'] = $searchBase }
    if ($global:cred) { $params['Credential'] = $global:cred }
    try {
        $days = [int]$numAdmin1Days.Value
        if ($days -eq 0) {
            $params['Filter'] = {adminCount -eq 1}
            $global:results = Get-ADUser @params | 
                              Select-Object Name, SamAccountName, adminCount, Enabled
            $htmlTitle = "AdminCount=1 Kullanıcılar"
        } else {
            $date = (Get-Date).AddDays(-$days)
            $params['Filter'] = {adminCount -eq 1 -and pwdLastSet -lt $date}
            $params['Properties'] += 'pwdLastSet'
            $global:results = Get-ADUser @params | 
                              Select-Object Name, SamAccountName, adminCount, @{Name='pwdLastSet';Expression={[datetime]::FromFileTime($_.pwdLastSet)}}, Enabled
            $htmlTitle = "Son $days Günde Şifre Değiştirmemiş AdminCount=1 Kullanıcılar"
        }
        $dataGrid.DataSource = $global:results
        if ($global:results) {
            $htmlContent = $global:results | ConvertTo-Html -Title "AD Sorgu Sonuçları" -PreContent "<h2>$htmlTitle</h2>" -Head $cssStyle
            $webBrowser.DocumentText = $htmlContent
        } else {
            $webBrowser.DocumentText = "<html><body><h2>Sonuç bulunamadı</h2></body></html>"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata: Sorgu başarısız! $_")
        $webBrowser.DocumentText = "<html><body><h2>Hata: Sorgu başarısız</h2></body></html>"
    }
})
$groupQueries.Controls.Add($btnAdmin1)

# NumericUpDown: AdminCount=0 sorgusu için (isteğe bağlı)
$lblAdmin0Days = New-Object System.Windows.Forms.Label
$lblAdmin0Days.Location = New-Object System.Drawing.Point(10,90)
$lblAdmin0Days.Size = New-Object System.Drawing.Size(150,20)
$lblAdmin0Days.Text = "AdminCount=0 için Gün:"
$groupQueries.Controls.Add($lblAdmin0Days)

$numAdmin0Days = New-Object System.Windows.Forms.NumericUpDown
$numAdmin0Days.Location = New-Object System.Drawing.Point(170,90)
$numAdmin0Days.Size = New-Object System.Drawing.Size(60,20)
$numAdmin0Days.Minimum = 0
$numAdmin0Days.Maximum = 365
$numAdmin0Days.Value = 0
$toolTip.SetToolTip($numAdmin0Days, "0 girerseniz gün filtresi uygulanmaz")
$groupQueries.Controls.Add($numAdmin0Days)

# Buton: AdminCount=0 olan kullanıcılar
$btnAdmin0 = New-Object System.Windows.Forms.Button
$btnAdmin0.Location = New-Object System.Drawing.Point(240,85)
$btnAdmin0.Size = New-Object System.Drawing.Size(150,30)
$btnAdmin0.Text = "AdminCount=0 Kullanıcılar"
$toolTip.SetToolTip($btnAdmin0, "AdminCount=0 veya tanımlı olmayan kullanıcıları listele")
$btnAdmin0.Add_Click({
    $domain = $txtDomain.Text
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen bir domain adı girin!")
        return
    }
    $searchBase = if ($comboOU.SelectedItem -and $comboOU.SelectedItem -ne "Tüm Domain (OU Yok)") { $comboOU.SelectedItem } else { $null }
    $params = @{ 
        Server = $domain
        Properties = @('Name', 'SamAccountName', 'adminCount', 'Enabled')
        ErrorAction = 'Stop'
    }
    if ($searchBase) { $params['SearchBase'] = $searchBase }
    if ($global:cred) { $params['Credential'] = $global:cred }
    try {
        $days = [int]$numAdmin0Days.Value
        if ($days -eq 0) {
            $params['Filter'] = {adminCount -eq 0 -or adminCount -notlike "*"}
            $global:results = Get-ADUser @params | 
                              Select-Object Name, SamAccountName, @{Name='adminCount';Expression={if($_.adminCount){$_.adminCount} else {0}}}, Enabled
            $htmlTitle = "AdminCount=0 Kullanıcılar"
        } else {
            $date = (Get-Date).AddDays(-$days)
            $params['Filter'] = "((adminCount -eq 0) -or (-not adminCount)) -and (pwdLastSet -lt $date)"
            $params['Properties'] += 'pwdLastSet'
            $global:results = Get-ADUser @params | 
                              Select-Object Name, SamAccountName, @{Name='adminCount';Expression={if($_.adminCount){$_.adminCount} else {0}}}, @{Name='pwdLastSet';Expression={[datetime]::FromFileTime($_.pwdLastSet)}}, Enabled
            $htmlTitle = "Son $days Günde Şifre Değiştirmemiş AdminCount=0 Kullanıcılar"
        }
        $dataGrid.DataSource = $global:results
        if ($global:results) {
            $htmlContent = $global:results | ConvertTo-Html -Title "AD Sorgu Sonuçları" -PreContent "<h2>$htmlTitle</h2>" -Head $cssStyle
            $webBrowser.DocumentText = $htmlContent
        } else {
            $webBrowser.DocumentText = "<html><body><h2>Sonuç bulunamadı</h2></body></html>"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata: Sorgu başarısız! $_")
        $webBrowser.DocumentText = "<html><body><h2>Hata: Sorgu başarısız</h2></body></html>"
    }
})
$groupQueries.Controls.Add($btnAdmin0)

# NumericUpDown: Login olmayan bilgisayarlar için gün sayısı
$lblCompDays = New-Object System.Windows.Forms.Label
$lblCompDays.Location = New-Object System.Drawing.Point(10,120)
$lblCompDays.Size = New-Object System.Drawing.Size(150,20)
$lblCompDays.Text = "Login Olmamış Bilgisayarlar:"
$groupQueries.Controls.Add($lblCompDays)

$numCompDays = New-Object System.Windows.Forms.NumericUpDown
$numCompDays.Location = New-Object System.Drawing.Point(170,120)
$numCompDays.Size = New-Object System.Drawing.Size(60,20)
$numCompDays.Minimum = 1
$numCompDays.Maximum = 365
$numCompDays.Value = 90
$toolTip.SetToolTip($numCompDays, "Son login süresini gün cinsinden belirtin")
$groupQueries.Controls.Add($numCompDays)

# Buton: Son X günde login olmamış bilgisayarlar
$btnInactiveComp = New-Object System.Windows.Forms.Button
$btnInactiveComp.Location = New-Object System.Drawing.Point(240,115)
$btnInactiveComp.Size = New-Object System.Drawing.Size(150,30)
$btnInactiveComp.Text = "Login Olmamış Bilgisayarlar"
$toolTip.SetToolTip($btnInactiveComp, "Belirtilen süre içinde login olmamış bilgisayarları listele")
$btnInactiveComp.Add_Click({
    $domain = $txtDomain.Text
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen bir domain adı girin!")
        return
    }
    $searchBase = if ($comboOU.SelectedItem -and $comboOU.SelectedItem -ne "Tüm Domain (OU Yok)") { $comboOU.SelectedItem } else { $null }
    $params = @{ 
        Filter = {lastLogonTimestamp -lt $date}
        Server = $domain
        Properties = @('Name', 'lastLogonTimestamp', 'Enabled')
        ErrorAction = 'Stop'
    }
    if ($searchBase) { $params['SearchBase'] = $searchBase }
    if ($global:cred) { $params['Credential'] = $global:cred }
    try {
        $days = [int]$numCompDays.Value
        $date = (Get-Date).AddDays(-$days)
        $global:results = Get-ADComputer @params | 
                          Select-Object Name, @{Name='lastLogonTimestamp';Expression={[datetime]::FromFileTime($_.lastLogonTimestamp)}}, Enabled
        $dataGrid.DataSource = $global:results
        if ($global:results) {
            $htmlContent = $global:results | ConvertTo-Html -Title "AD Sorgu Sonuçları" -PreContent "<h2>Son $days Günde Login Olmamış Bilgisayarlar</h2>" -Head $cssStyle
            $webBrowser.DocumentText = $htmlContent
        } else {
            $webBrowser.DocumentText = "<html><body><h2>Sonuç bulunamadı</h2></body></html>"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata: Sorgu başarısız! $_")
        $webBrowser.DocumentText = "<html><body><h2>Hata: Sorgu başarısız</h2></body></html>"
    }
})
$groupQueries.Controls.Add($btnInactiveComp)

# GroupBox: İşlemler
$groupActions = New-Object System.Windows.Forms.GroupBox
$groupActions.Location = New-Object System.Drawing.Point(10,350)
$groupActions.Size = New-Object System.Drawing.Size(1160,80)
$groupActions.Text = "Bilgisayar İşlemleri"
$form.Controls.Add($groupActions)

# Buton: Bilgisayarı Kapat
$btnShutdown = New-Object System.Windows.Forms.Button
$btnShutdown.Location = New-Object System.Drawing.Point(10,30)
$btnShutdown.Size = New-Object System.Drawing.Size(150,30)
$btnShutdown.Text = "Bilgisayarı Kapat"
$toolTip.SetToolTip($btnShutdown, "Seçilen bilgisayarı kapat")
$btnShutdown.Add_Click({
    if (-not $comboComputers.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen listeden bir bilgisayar seçin!")
        return
    }
    $computer = $comboComputers.SelectedItem
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$computer bilgisayarı kapatılacak. Devam etmek istiyor musunuz?",
        "Onay", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne 'Yes') { return }
    
    try {
        if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
            throw "Bilgisayar çevrimdışı veya ulaşılamıyor"
        }
        $params = @{ ComputerName = $computer; Force = $true; ErrorAction = 'Stop' }
        if ($global:cred) { $params['Credential'] = $global:cred }
        Stop-Computer @params
        [System.Windows.Forms.MessageBox]::Show("$computer başarıyla kapatıldı.", "Başarılı", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata ($computer): $_", "Hata", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$groupActions.Controls.Add($btnShutdown)

# NetBIOS için Bilgisayar İsmi Girişi
$lblNetBIOSComputer = New-Object System.Windows.Forms.Label
$lblNetBIOSComputer.Location = New-Object System.Drawing.Point(170,30)
$lblNetBIOSComputer.Size = New-Object System.Drawing.Size(150,20)
$lblNetBIOSComputer.Text = "NetBIOS Bilgisayar:"
$groupActions.Controls.Add($lblNetBIOSComputer)

$txtNetBIOSComputer = New-Object System.Windows.Forms.TextBox
$txtNetBIOSComputer.Location = New-Object System.Drawing.Point(330,30)
$txtNetBIOSComputer.Size = New-Object System.Drawing.Size(200,20)
$txtNetBIOSComputer.Text = ""
$toolTip.SetToolTip($txtNetBIOSComputer, "NetBIOS'u devre dışı bırakmak için bilgisayar ismi girin veya listeden seçin")
$groupActions.Controls.Add($txtNetBIOSComputer)

# Buton: NetBIOS over TCP/IP'yi Devre Dışı Bırak
$btnDisableNetBIOS = New-Object System.Windows.Forms.Button
$btnDisableNetBIOS.Location = New-Object System.Drawing.Point(540,25)
$btnDisableNetBIOS.Size = New-Object System.Drawing.Size(150,30)
$btnDisableNetBIOS.Text = "NetBIOS'u Devre Dışı Bırak"
$toolTip.SetToolTip($btnDisableNetBIOS, "Girilen veya seçilen bilgisayarda NetBIOS'u devre dışı bırak")
$btnDisableNetBIOS.Add_Click({
    $computer = if ($txtNetBIOSComputer.Text) { $txtNetBIOSComputer.Text } else { $comboComputers.SelectedItem }
    if (-not $computer) {
        [System.Windows.Forms.MessageBox]::Show("Lütfen listeden bir bilgisayar seçin veya bilgisayar ismi girin!")
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$computer bilgisayarında NetBIOS over TCP/IP devre dışı bırakılacak.`nDevam etmek istiyor musunuz?`nNot: Bu işlem bazı eski uygulamaları etkileyebilir.",
        "Onay", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne 'Yes') { return }
    
    try {
        if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
            throw "Bilgisayar çevrimdışı veya ulaşılamıyor"
        }
        $sessionParams = @{ ComputerName = $computer; ErrorAction = 'Stop' }
        if ($global:cred) { $sessionParams['Credential'] = $global:cred }
        $session = New-PSSession @sessionParams
        Invoke-Command -Session $session -ScriptBlock {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces"
            $interfaces = Get-ChildItem -Path $regPath -ErrorAction Stop
            if (-not $interfaces) { throw "Ağ arabirimi bulunamadı" }
            foreach ($interface in $interfaces) {
                Set-ItemProperty -Path "$regPath\$($interface.PSChildName)" -Name "NetbiosOptions" -Value 2 -ErrorAction Stop
            }
        }
        Remove-PSSession $session
        [System.Windows.Forms.MessageBox]::Show("$computer üzerinde NetBIOS başarıyla devre dışı bırakıldı.", "Başarılı", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hata ($computer): $_", "Hata", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$groupActions.Controls.Add($btnDisableNetBIOS)

# DataGridView (sonuçları göstermek için)
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10,440)
$dataGrid.Size = New-Object System.Drawing.Size(570,300) # Genişlik daraltıldı
$dataGrid.ReadOnly = $true
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.MultiSelect = $true
$dataGrid.RowHeadersVisible = $true
$toolTip.SetToolTip($dataGrid, "Sorgu sonuçlarını görüntüleyin")
$form.Controls.Add($dataGrid)

# WebBrowser (HTML önizlemesi için)
$webBrowser = New-Object System.Windows.Forms.WebBrowser
$webBrowser.Location = New-Object System.Drawing.Point(590,440) # Sağ tarafa taşındı
$webBrowser.Size = New-Object System.Drawing.Size(580,300)
$webBrowser.IsWebBrowserContextMenuEnabled = $false
$toolTip.SetToolTip($webBrowser, "Sorgu sonuçlarının HTML önizlemesi")
$form.Controls.Add($webBrowser)

# Sonuç verisi (global değişken)
$global:results = $null
$global:cred = $null

# CSS stili (HTML için)
$cssStyle = "<style>body {font-family: Arial, sans-serif;} table {border-collapse: collapse; width: 100%;} th, td {border: 1px solid #ddd; padding: 8px; text-align: left;} th {background-color: #f2f2f2; color: black;} tr:nth-child(even) {background-color: #f9f9f9;} tr:hover {background-color: #ddd;}</style>"

# Buton: CSV Olarak Kaydet
$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Location = New-Object System.Drawing.Point(10,750)
$btnCsv.Size = New-Object System.Drawing.Size(150,30)
$btnCsv.Text = "CSV Kaydet"
$toolTip.SetToolTip($btnCsv, "Sorgu sonuçlarını CSV dosyasına kaydet")
$btnCsv.Add_Click({
    if ($global:results) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
        if ($saveDialog.ShowDialog() -eq "OK") {
            $global:results | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("CSV kaydedildi: " + $saveDialog.FileName)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Önce bir sorgu çalıştırın!")
    }
})
$form.Controls.Add($btnCsv)

# Buton: HTML Olarak Kaydet
$btnHtml = New-Object System.Windows.Forms.Button
$btnHtml.Location = New-Object System.Drawing.Point(170,750)
$btnHtml.Size = New-Object System.Drawing.Size(150,30)
$btnHtml.Text = "HTML Kaydet"
$toolTip.SetToolTip($btnHtml, "Sorgu sonuçlarını HTML dosyasına kaydet")
$btnHtml.Add_Click({
    if ($global:results) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "HTML Files (*.html)|*.html"
        if ($saveDialog.ShowDialog() -eq "OK") {
            $htmlContent = $global:results | ConvertTo-Html -Title "AD Sorgu Sonuçları" -Head $cssStyle
            $htmlContent | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("HTML kaydedildi: " + $saveDialog.FileName)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Önce bir sorgu çalıştırın!")
    }
})
$form.Controls.Add($btnHtml)

# Formu göster
$form.ShowDialog()