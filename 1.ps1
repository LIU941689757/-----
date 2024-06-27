Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 创建表单
$form = New-Object System.Windows.Forms.Form
$form.Text = "出勤腿勤记录"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = "CenterScreen"

# 实时日期和时间标签
$datetimeLabel = New-Object System.Windows.Forms.Label
$datetimeLabel.Location = New-Object System.Drawing.Point(10,10)
$datetimeLabel.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($datetimeLabel)

# 出勤按钮
$checkInButton = New-Object System.Windows.Forms.Button
$checkInButton.Text = "出勤"
$checkInButton.Location = New-Object System.Drawing.Point(50,50)
$checkInButton.Size = New-Object System.Drawing.Size(75,23)
$form.Controls.Add($checkInButton)

# 退勤按钮
$checkOutButton = New-Object System.Windows.Forms.Button
$checkOutButton.Text = "退勤"
$checkOutButton.Location = New-Object System.Drawing.Point(150,50)
$checkOutButton.Size = New-Object System.Drawing.Size(75,23)
$form.Controls.Add($checkOutButton)

# 更新实时日期和时间
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    $datetimeLabel.Text = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
})
$timer.Start()

# 出勤按钮点击事件
$checkInButton.Add_Click({
    $currentTime = Get-Date
    $dateString = $currentTime.ToString("yyyy-MM-dd")
    $filePath = "$dateString-出勤时间.txt"
    $currentTime | Out-File -FilePath $filePath
    [System.Windows.Forms.MessageBox]::Show("出勤时间记录成功！")
})

# 退勤按钮点击事件
$checkOutButton.Add_Click({
    $currentDate = Get-Date
    $dateString = $currentDate.ToString("yyyy-MM-dd")
    $filePathIn = "$dateString-出勤时间.txt"
    $filePathOut = "$dateString-腿勤时间.txt"

    if (Test-Path $filePathIn) {
        $checkInTime = Get-Content -Path $filePathIn | Out-String | Get-Date
        $checkOutTime = Get-Date
        $workDuration = $checkOutTime - $checkInTime
        $output = "出勤时间: $checkInTime`n退勤时间: $checkOutTime`n工作时长: $($workDuration.Hours)小时$($workDuration.Minutes)分钟$($workDuration.Seconds)秒"
        $output | Out-File -FilePath $filePathOut
        [System.Windows.Forms.MessageBox]::Show("退勤时间记录成功！")
    } else {
        [System.Windows.Forms.MessageBox]::Show("未找到出勤记录，请先记录出勤时间。")
    }
})

# 运行表单
[void]$form.ShowDialog()
