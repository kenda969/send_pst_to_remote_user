Add-Type -assembly System.Windows.Forms

$main_form = New-Object System.Windows.Forms.Form
$Label = New-Object System.Windows.Forms.Label
$TextBox = New-Object System.Windows.Forms.TextBox
$OKButton = New-Object System.Windows.Forms.Button

$main_form.Text = 'Копирование архива'
$main_form.Width = 270
$main_form.Height = 260
$main_form.AutoSize = $True


$Label.Text = "Введите пользователей для копирования"
$Label.Location  = New-Object System.Drawing.Point(20,10)
$Label.Size = New-Object System.Drawing.Size(225,20)
$main_form.Controls.Add($Label)


$TextBox.Location  = New-Object System.Drawing.Point(20,30)
$TextBox.Size = New-Object System.Drawing.Size(220,150)
$TextBox.MultiLine = $True
$TextBox.Text = 'Пример: host-name.'
$TextBox.Place
$main_form.Controls.Add($TextBox)


$OKButton.Location = New-Object System.Drawing.Point(90,190)
$OKButton.Size = New-Object System.Drawing.Size(80,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$main_form.Controls.Add($OKButton)

$result = $main_form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]:: OK){

    $host_name = $textBox.Text.Split(';',100).Trim()
    $host_name

    $openFileDialog = New-Object windows.forms.openfiledialog

    $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory() 
    $openFileDialog.title = "Выберите архив для копирования(формат PST)"   
    $openFileDialog.filter = "All files (*.*)| *.*"
    $openFileDialog.filter = "PublishSettings Files|*.publishsettings|All Files|*.*" 
    $openFileDialog.ShowHelp = $True
     
    Write-Host "Select Downloaded Settings File... (see FileOpen Dialog)" -ForegroundColor Green

    $result = $openFileDialog.ShowDialog()

    if($result -eq "OK")    {   
    Add-Type -assembly System.Windows.Forms

$archive_name_form = New-Object System.Windows.Forms.Form
$Label = New-Object System.Windows.Forms.Label
$TextBox = New-Object System.Windows.Forms.TextBox
$OKButton = New-Object System.Windows.Forms.Button

$archive_name_form.Text = 'Название архива'
$archive_name_form.Width = 270
$archive_name_form.Height = 260
$archive_name_form.AutoSize = $false


$Label.Text = "Введите название архива"
$Label.Location  = New-Object System.Drawing.Point(20,10)
$Label.Size = New-Object System.Drawing.Size(225,20)
$archive_name_form.Controls.Add($Label)


$TextBox.Location  = New-Object System.Drawing.Point(20,30)
$TextBox.Size = New-Object System.Drawing.Size(220,150)
$TextBox.MultiLine = $False
$TextBox.Text = 'Archive_.pst'
$TextBox.Place
$archive_name_form.Controls.Add($TextBox)


$OKButton.Location = New-Object System.Drawing.Point(90,190)
$OKButton.Size = New-Object System.Drawing.Size(80,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$archive_name_form.Controls.Add($OKButton) 

$result = $archive_name_form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]:: OK){

    $archive_name = $textBox.Text

   $z = 0

foreach($j in $host_name){
$h_name = $j.Split(';')[0]
$h_name

 $server_disk = invoke-command -computername $h_name {Get-WmiObject Win32_LogicalDisk | where {$_.DriveType  -eq '3'}}

    foreach($i in  $server_disk){

    $user_name = Get-WMIObject -Class Win32_ComputerSystem -Computer  $h_name |Select username
    $user_name = [string]$user_name
    $user_name = $user_name.trimStart('@{username=VELESSTROY')
    $user_name = $user_name.trimEnd('}')

        $size = $($i.size/1GB).ToString().Split(',')[0]
        $free = $($i.freeSpace/1GB).ToString().Split(',')[0]
       
        if($($z % 2) -eq 0 ){
            Write-Host -ForegroundColor Yellow $h_name.ToString() "Disk:=" $i.DeviceID[0],  'Size(GB):=' $($Size), 'Freesize(GB):=' $($free) 
        }
        else{
            Write-Host $h_name.ToString() "Disk:=" $i.DeviceID[0],  'Size(GB):=' $($Size), 'Freesize(GB):=' $($free) 
        }
    
        $z++
   
        if($($i.freeSpace/1GB) -ge 20){
            $path_to_exe = 'C:\archive_to_remote_usr\add_pst.exe'
            if($i.DeviceID[0] -eq 'C'){

            $path_to_dir = '\\' + $h_name + '\'+ $i.DeviceID[0] +'$\Users'+ $user_name + '\Documents\Файлы Outlook\'
            
                if(!(Test-Path $path_to_dir)){
                    New-Item -ItemType Directory -Force -Path $path_to_dir
                    Copy-item $OpenFileDialog.filename $path_to_dir$archive_name -force
                    New-Item -Path $path_to_dir'path_pst.txt' -ItemType File
                    Add-content -Path  $path_to_dir'path_pst.txt' $path_to_dir$archive_name
                    Copy-item $path_to_exe $path_to_dir'add_pst.exe' -force
                    Write-Host $user_name
                    break
                    }else{
                    Copy-item $OpenFileDialog.filename $path_to_dir$archive_name -force
                    New-Item -Path $path_to_dir'path_pst.txt' -ItemType File
                    Add-content -Path  $path_to_dir'path_pst.txt' $path_to_dir$archive_name
                    if(!(Test-Path -Path $path_to_dir'add_pst.exe')){
                            Copy-item $path_to_exe $path_to_dir'add_pst.exe' -force

                    }
                    Write-Host $user_name
                    break
                    }
            }

            if($i.DeviceID[0] -eq 'D'){

                $path_to_dir = '\\' + $h_name + '\' + $i.DeviceID[0] + '$\файлы Outlook\'

                if(!(test-path $path_to_dir)){

                    New-Item -ItemType Directory -Force -Path $path_to_dir
                    Copy-item $OpenFileDialog.filename $path_to_dir$archive_name -force
                    New-Item -Path $path_to_dir'path_pst.txt' -ItemType File
                    Add-content -Path  $path_to_dir'path_pst.txt' $path_to_dir$archive_name
                    Copy-item $path_to_exe $path_to_dir'add_pst.exe' -force

                    Write-Host $user_name
                    break
                }else{
                      Copy-item $OpenFileDialog.filename $path_to_dir$archive_name -force
                      New-Item -Path $path_to_dir'path_pst.txt' -ItemType File
                      Add-content -Path  $path_to_dir'path_pst.txt' $path_to_dir$archive_name
                      if(!(Test-Path -Path $path_to_dir'add_pst.exe')){
                          Copy-item $path_to_exe $path_to_dir'add_pst.exe' -force

                      }
                      Write-Host $user_name
                      break
                }
               }
        }else{
            Write-Host 'Нет свободного места на диске' $i.DeviceID[0] 'компьютер' $h_name
        }
    }
}
    }else{Write-Host "Отмена!" -ForegroundColor Yellow}
    }else { Write-Host "Ничего не выбрано!" -ForegroundColor Yellow} 
 }

