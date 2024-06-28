#######################################################
#
#   VERSION：0.2
#   作者：LIU YUXI
#   更新日期：2024.06.28
#   备考：
#
#######################################################


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 创建表单
$form = New-Object System.Windows.Forms.Form
$form.Text = "出勤退勤记录"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = "CenterScreen"


# 实时日期和时间标签
$datetimeLabel = New-Object System.Windows.Forms.Label
$datetimeLabel.Location = New-Object System.Drawing.Point(40,10)
#字号
$datetimeLabel.Font = New-Object System.Drawing.Font("Arial", 15, [System.Drawing.FontStyle]::Regular)
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

# 计算总工作时长按钮
$totalDurationButton = New-Object System.Windows.Forms.Button
$totalDurationButton.Text = "计算总工作时长"
$totalDurationButton.Location = New-Object System.Drawing.Point(50,90)
$totalDurationButton.Size = New-Object System.Drawing.Size(175,23)
$form.Controls.Add($totalDurationButton)

# 更新实时日期和时间
$timer = New-Object System.Windows.Forms.Timer  # 创建一个新的定时器对象
# 设置定时器的间隔时间为 1000 毫秒（即 1 秒）
$timer.Interval = 1000
# 添加一个 Tick 事件处理程序，每当定时器间隔时间结束时，执行下面的代码块
$timer.Add_Tick({
    # 更新 $datetimeLabel 控件的文本属性，显示当前日期和时间
    $datetimeLabel.Text = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
})
# 启动定时器
$timer.Start()


# 出勤按钮点击事件
$checkInButton.Add_Click({
    $currentTime = Get-Date
    $dateString = $currentTime.ToString("yyyy-MM-dd")
    $filePath = "$dateString-出勤.txt"
    $currentTime | Out-File -FilePath $filePath
    [System.Windows.Forms.MessageBox]::Show("出勤时间记录成功！")
})

# 退勤按钮点击事件
$checkOutButton.Add_Click({
    $currentDate = Get-Date
    $dateString = $currentDate.ToString("yyyy-MM-dd")
    $filePathIn = "$dateString-出勤.txt"
    $filePathOut = "$dateString-退勤.txt"

    if (Test-Path $filePathIn) {
        # 从文件中获取内容并转换为日期对象
        $checkInTime = Get-Content -Path $filePathIn | Out-String | Get-Date
        $checkOutTime = Get-Date
        $workDuration = $checkOutTime - $checkInTime
        $output = "出勤时间: $checkInTime`n退勤时间: $checkOutTime`n工作时长: $($workDuration.Hours)小时$($workDuration.Minutes)分钟$($workDuration.Seconds)秒"
        $output | Out-File -FilePath $filePathOut
        [System.Windows.Forms.MessageBox]::Show("退勤时间记录成功！`n$($output)")
    } else {
        [System.Windows.Forms.MessageBox]::Show("未找到出勤记录，请先记录出勤时间。")
    }
})

# 计算总工作时长按钮点击事件
$totalDurationButton.Add_Click({
    $totalDuration = [TimeSpan]::Zero
    $files = Get-ChildItem -Filter "*-退勤.txt"
    
# 初始化总工作时长变量
$totalDuration = [System.TimeSpan]::Zero

# 遍历文件列表中的每个文件
foreach ($file in $files) {
    # 获取文件内容
    $content = Get-Content -Path $file.FullName
    
    # 遍历文件内容的每一行
    foreach ($line in $content) {
        # 检查每一行是否匹配工作时长的格式
        if ($line -match "工作时长: (\d+)小时(\d+)分钟(\d+)秒") {
            # 提取小时、分钟和秒
            $hours = [int]$matches[1]
            $minutes = [int]$matches[2]
            $seconds = [int]$matches[3]
            
            # 创建时间间隔对象并累加到总时长中
            $totalDuration += New-TimeSpan -Hours $hours -Minutes $minutes -Seconds $seconds
        }
    }
}

    
    $totalOutput = "总工作时长: $($totalDuration.Days * 24 + $totalDuration.Hours)小时$($totalDuration.Minutes)分钟$($totalDuration.Seconds)秒"
    [System.Windows.Forms.MessageBox]::Show($totalOutput)
})

# 运行表单
[void]$form.ShowDialog()
