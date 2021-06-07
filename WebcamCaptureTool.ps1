# ======================================================================
#
# Webcam Capture Tool
#
# Copyright (C) 2021 tag. All rights reserved.
#
# ======================================================================

Set-StrictMode -Version Latest

# デバッグモード
$DebugPreference = "SilentlyContinue"
#$DebugPreference = "Continue"

# スクリプト情報の取得＋フォルダへの移動
$ScriptDir = ""
$ScriptName = ""
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $ScriptPath = $MyInvocation.MyCommand.Definition
    $ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 
    $ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)
} else {
    $ScriptPath = [Environment]::GetCommandLineArgs()[0]
    $ScriptDir = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
    $ScriptName = [System.IO.Path]::GetFileNameWithoutExtension([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptDir) {
        $ScriptDir = "." 
    }
}
Write-Debug "> Script Path: $ScriptPath"
Write-Debug "> Script Directory: $ScriptDir"
Write-Debug "> Script Name: $ScriptName"
Set-Location $ScriptDir

# 多重起動禁止
Write-Debug "多重起動チェック"
$Mutex = New-Object System.Threading.Mutex -ArgumentList $false, "Global¥$ScriptName"
try {
    if (-not $Mutex.WaitOne(0, $false)) {
        Write-Debug "> すでに起動済みです。終了します。"
        $Mutex.Close()
        exit(0)
    }
} catch [System.Threading.AbandonedMutexException] {
    Write-Debug "> 前回の実行は強制終了しています。"
}

# ツール名
$AppName = "Webcam Capture Tool"

Write-Debug ""
Write-Debug "======================================================================"
Write-Debug ""
Write-Debug $AppName
Write-Debug ""
Write-Debug "Copyright (C) 2021 tag. All rights reserved."
Write-Debug ""
Write-Debug "======================================================================"
Write-Debug ""

# 設定ファイルの読み込み
$ConfigFile = Join-Path -Path $ScriptDir -ChildPath "$ScriptName.ini"
if (Test-Path $ConfigFile) {
    $Params = @{}
    Get-Content $ConfigFile | % { $Params += ConvertFrom-StringData $_ }
}

# 画像保存先
$SaveDirectory = $ScriptDir
if ($Params.SAVE_DIRECTORY -ne $null) {
    $SaveDirectory = $Params.SAVE_DIRECTORY
}
Write-Debug "> Save Directory: $SaveDirectory"

Write-Debug "----------------------------------------------------------------------"
Write-Debug "動作環境"
$Version = $PSVersionTable.PSVersion
Write-Debug "> PowerShell: $Version"
$Architecture = "Unknown"
if ([System.Environment]::Is64BitProcess) {
    $Architecture = "x64"
} else {
    $Architecture = "x86"
}
Write-Debug "> Process: $Architecture"
$LibDir = Join-Path -Path $ScriptDir -ChildPath $Architecture


Write-Debug "----------------------------------------------------------------------"
Write-Debug "OpenCvSharpExtern.dll のコピー"
$OpenCvSharpExternDLLFileName = "OpenCvSharpExtern.dll"
$OpenCvSharpExternDLLFilePath = $(Join-Path -Path $LibDir -ChildPath $OpenCvSharpExternDLLFileName)
if (-not (Test-Path $(Join-Path -Path $ScriptDir -ChildPath "$OpenCvSharpExternDLLFileName"))) {
    if (Test-Path $OpenCvSharpExternDLLFilePath) {
        Copy-Item $OpenCvSharpExternDLLFilePath $(Join-Path -Path $ScriptDir -ChildPath "$OpenCvSharpExternDLLFileName") -Force
    } else {
        Write-Host "$OpenCvSharpExternDLLFileName が見つかりません。"
        [System.Windows.Forms.MessageBox]::Show("$OpenCvSharpExternDLLFileName が見つかりません。", "エラー") 
        exit(-1)
    }
}

Write-Debug "----------------------------------------------------------------------"
Write-Debug "アセンブリのロード"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Debug "----------------------------------------------------------------------"
Write-Debug "OpenCvSharp のロード"
$OpenCvSharpDLL = Join-Path -Path $ScriptDir -ChildPath "OpenCvSharp.dll"
if (Test-Path $OpenCvSharpDLL) {
    $Res = [System.Reflection.Assembly]::LoadFrom($OpenCvSharpDLL)
    Write-Debug "> $Res"
} else {
    Write-Host "$OpenCvSharpDLL が見つかりません。"
    [System.Windows.Forms.MessageBox]::Show("$OpenCvSharpDLL が見つかりません。", "エラー") 
    exit(-1)
}

$OpenCvSharpExtDLL = Join-Path -Path $ScriptDir -ChildPath "OpenCvSharp.Extensions.dll"
if (Test-Path $OpenCvSharpExtDLL) {
    $Res = [System.Reflection.Assembly]::LoadFrom($OpenCvSharpExtDLL)
    Write-Debug "> $Res"
} else {
    Write-Host "$OpenCvSharpExtDLL が見つかりません。"
    [System.Windows.Forms.MessageBox]::Show("$OpenCvSharpExtDLL が見つかりません。", "エラー") 
    exit(-1)
}

Write-Debug "----------------------------------------------------------------------"
Write-Debug "アセンブリロード状況"
$Res = [System.AppDomain]::CurrentDomain.GetAssemblies() | % { $_.GetName().Name }
foreach ($record in $Res) {
    Write-Debug "> $record"
}

Write-Debug "----------------------------------------------------------------------"
Write-Debug "カメラ関連関数の準備"

# 重複保存禁止フラグ
$global:AlreadySaved = $false

$LoadDialogWidth = 320
$LoadDialogHeight = 64
 
$global:LoadForm = New-Object System.Windows.Forms.Form
$global:LoadForm.ClientSize = New-Object System.Drawing.Size($LoadDialogWidth, $LoadDialogHeight)
$global:LoadForm.StartPosition = "CenterScreen"
$global:LoadForm.AutoSize = $false
$global:LoadForm.FormBorderStyle = "None"
$global:LoadForm.MaximizeBox = $false
$global:LoadForm.MinimizeBox = $false
$global:LoadForm.Topmost = $true

$global:LoadMessageLabel = New-Object System.Windows.Forms.Label
$global:LoadMessageLabel.Location = New-Object System.Drawing.Point(0, 0)
$global:LoadMessageLabel.Size = New-Object System.Drawing.Size($LoadDialogWidth, $LoadDialogHeight)
$global:LoadMessageLabel.Text = "起動中..."
$global:LoadMessageLabel.Font = [System.Drawing.Font]::new("Meiryo UI", 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, 128)
$global:LoadMessageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$global:LoadForm.Controls.Add($global:LoadMessageLabel)

[void] $global:LoadForm.Show()

$CaptureWidth = [int]1280
$CaptureHeight = [int]720
if ($Params.CAPTURE_WIDTH -ne $null) {
    $CaptureWidth = [int]$Params.CAPTURE_WIDTH
}
Write-Debug "> Capture Width: $CaptureWidth"
if ($Params.CAPTURE_HEIGHT -ne $null) {
    $CaptureHeight = [int]$Params.CAPTURE_HEIGHT
}
Write-Debug "> Capture Height: $CaptureHeight"

$global:Capture = [OpenCvSharp.VideoCapture]::new()
[void]$global:Capture.Open(0)
[void]$global:Capture.Set([OpenCvSharp.VideoCaptureProperties]::FrameWidth, $CaptureWidth)
[void]$global:Capture.Set([OpenCvSharp.VideoCaptureProperties]::FrameHeight, $CaptureHeight)

$global:LoadForm.Close()

if (-not $global:Capture.IsOpened()) {
    Write-Host "カメラのオープンに失敗しました。"
    [System.Windows.Forms.MessageBox]::Show("カメラのオープンに失敗しました。カメラを使用しているアプリケーションを使用している場合には、そのアプリケーションを終了した後、再試行してください。", "エラー") 
    FinalizeCamera
    exit(-1)
}

function FinalizeCamera() {
    Write-Debug "カメラの終了"
    if ($global:Capture.IsOpened()) {
        $global:Capture.Release()
        $global:Capture.Dispose()
    }
}

function Capture() {
    Write-Debug "カメラからフレームを取得"
    try {
        $Frame = [OpenCvSharp.Mat]::new()
        $global:Capture.Read($Frame)
        $global:Label.Text = "撮影に成功しました。表示されている画像を保存する場合には [保存] をクリックしてください。"
        if ($Frame.Empty()) {
            Write-Host "フレームの取得に失敗"
            $global:Label.Text = "エラー: フレームの取得に失敗しました。"
            FinalizeCamera
            exit(-1)
        }
        $Bitmap = [OpenCvSharp.Extensions.BitmapConverter]::ToBitmap($Frame)
        $global:PictureBox.Image = $Bitmap
        $global:AlreadySaved = $false
    } catch {
        Write-Debug $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("カメラのオープンに失敗しました。カメラを使用しているアプリケーションを使用している場合には、そのアプリケーションを終了した後、再試行してください。", "エラー") 
    }
}

function SaveAsFile() {
    Write-Debug "ファイルで保存"

    if ($global:AlreadySaved) {
        $global:Label.Text = "すでに保存済みのデータであるため、保存されませんでした。"
        return
    }

    $Bitmap = $global:PictureBox.Image

    $FilePath = Join-Path -Path $SaveDirectory -ChildPath $(GenerateFilename)
    Write-Debug "> $FilePath"

    $Bitmap.Save($FilePath)

    if (Test-Path $FilePath) {
        $global:Label.Text = "保存に成功しました。 (保存先: $FilePath)"
        $global:Alreadysaved = $true
    } else {
        $global:Label.Text = "エラー: 保存に失敗しました。 (保存先: $FilePath)"
        [System.Windows.Forms.MessageBox]::Show("保存に失敗しました。撮影が完了しているか、保存先にアクセスできるか確認してください。", "エラー") 
    }
}

function GenerateFilename() {
    $Filename = "Capture" + "_" + $env:USERNAME + "_" + $env:COMPUTERNAME+ "_" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".png"
    return [string]$Filename
}

Write-Debug "----------------------------------------------------------------------"
Write-Debug "ウインドウの生成"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles()

$Margin = [int]24
$ButtonWidth = [int]150
$ButtonHeight = [int]32
$LabelWidth = [int]($CaptureWidth - $Margin * 2)
$LabelHeight = [int]24

# フォーム
$global:MainForm = New-Object System.Windows.Forms.Form
$global:MainForm.ClientSize = New-Object System.Drawing.Size([int]($CaptureWidth + $Margin * 2), [int]($CaptureHeight + $ButtonHeight + $Margin * 3))
$global:MainForm.StartPosition = "CenterScreen"
$global:MainForm.AutoSize = $false
$global:MainForm.FormBorderStyle = "FixedSingle"
$global:MainForm.MaximizeBox = $false
$global:MainForm.MinimizeBox = $false
$global:MainForm.Text = "$AppName ($env:USERNAME@$env:COMPUTERNAME)"

# メニューストリップ
$MenuStrip = New-Object System.Windows.Forms.MenuStrip
$global:MainForm.MainMenuStrip = $MenuStrip
$global:MainForm.Controls.Add($MenuStrip)

# メニュー - 撮影
$CaptureMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$CaptureMenuItem.Text = "撮影(&W)"
$CaptureMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::W
$CaptureMenuItem.Add_Click({
    Capture
})
[void]$MenuStrip.Items.Add($CaptureMenuItem)

# メニュー - 保存
$SaveAsFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$SaveAsFileMenuItem.Text = "保存(&S)"
$SaveAsFileMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::S
$SaveAsFileMenuItem.Add_Click({
    SaveAsFile
})
[void]$MenuStrip.Items.Add($SaveAsFileMenuItem)

# メニュー - 終了
$QuitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$QuitMenuItem.Text = "終了(&Q)"
$QuitMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::Q
$QuitMenuItem.Add_Click({
    $global:MainForm.Close()
})
[void]$MenuStrip.Items.Add($QuitMenuItem)

# メニュー - アプリについて
$AboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutMenuItem.Text = "アプリについて(&A)"
$AboutMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::A
$AboutMenuItem.Add_Click({ 
    
    $AboutFormMargin = [int]24
    $AboutFormWidth = [int]320
    $AboutFormHeight = [int]220
    $AboutLabelHeight = [int]36

    $AboutForm = New-Object System.Windows.Forms.Form
    $AboutForm.ClientSize = New-Object System.Drawing.Size($AboutFormWidth, $AboutFormHeight)
    $AboutForm.StartPosition = "CenterScreen"
    $AboutForm.AutoSize = $false
    $AboutForm.FormBorderStyle = "FixedSingle"
    $AboutForm.MaximizeBox = $false
    $AboutForm.MinimizeBox = $false
    $AboutForm.Text = "アプリについて"
    $AboutForm.Owner = $global:MainForm

    # アプリ名
    $AboutAppNameLabel = New-Object System.Windows.Forms.Label
    $AboutAppNameLabel.Location = New-Object System.Drawing.Point([int]($AboutFormMargin), [int]($AboutFormMargin))
    $AboutAppNameLabel.Size = New-Object System.Drawing.Size([int]($AboutFormWidth - $AboutFormMargin * 2), [int]($AboutLabelHeight))
    $AboutAppNameLabel.Text = $AppName
    $AboutAppNameLabel.Font = [System.Drawing.Font]::new("MS UI Gothic", 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, 128)
    $AboutAppNameLabel.TextAlign = "MiddleCenter"
    $AboutForm.Controls.Add($AboutAppNameLabel)

    # コピーライト
    $AboutCopyrightLabel = New-Object System.Windows.Forms.Label
    $AboutCopyrightLabel.Location = New-Object System.Drawing.Point([int]($AboutFormMargin), [int]($AboutFormMargin + $AboutLabelHeight))
    $AboutCopyrightLabel.Size = New-Object System.Drawing.Size([int]($AboutFormWidth - $AboutFormMargin * 2), [int]($AboutLabelHeight))
    $AboutCopyrightLabel.Text = "Copyright (c) 2021 tag. All rights reserved."
    $AboutCopyrightLabel.Font = [System.Drawing.Font]::new("MS UI Gothic", 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 128)
    $AboutCopyrightLabel.TextAlign = "MiddleCenter"
    $AboutForm.Controls.Add($AboutCopyrightLabel)

    # 環境変数
    $AboutCopyrightLabel = New-Object System.Windows.Forms.Label
    $AboutCopyrightLabel.Location = New-Object System.Drawing.Point([int]($AboutFormMargin), [int]($AboutFormMargin + $AboutLabelHeight * 2))
    $AboutCopyrightLabel.Size = New-Object System.Drawing.Size([int]($AboutFormWidth - $AboutFormMargin * 2), [int]($AboutLabelHeight * 3))
    $AboutCopyrightLabel.Text = "Powered by`n`nPowerShell $Version ($Architecture)`nOpenCvSharp"
    $AboutCopyrightLabel.Font = [System.Drawing.Font]::new("MS UI Gothic", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 128)
    $AboutCopyrightLabel.TextAlign = "MiddleCenter"
    $AboutForm.Controls.Add($AboutCopyrightLabel)

    [void] $AboutForm.Show()
})
[void]$MenuStrip.Items.Add($AboutMenuItem)

# ピクチャーボックス
$global:PictureBox = New-Object System.Windows.Forms.PictureBox
$global:PictureBox.Location = New-Object System.Drawing.Point($Margin, $Margin)
$global:PictureBox.Size = New-Object System.Drawing.Size($CaptureWidth, $CaptureHeight)
$global:MainForm.Controls.Add($PictureBox)

# ラベルを表示
$global:Label = New-Object System.Windows.Forms.Label
$global:Label.Location = New-Object System.Drawing.Point([int]($Margin), [int]($CaptureHeight + $Margin * 2))
$global:Label.Size = New-Object System.Drawing.Size($LabelWidth, $LabelHeight)
$global:Label.Text = "[撮影] をクリックしてください。"
$global:Label.TextAlign = "MiddleLeft"
$global:MainForm.Controls.Add($global:Label)

# フォームのアクティベート
$global:MainForm.Add_Shown({
    #Capture
    $global:MainForm.Activate()
})
[void] $global:MainForm.ShowDialog()

FinalizeCamera

$Mutex.ReleaseMutex()
$Mutex.Close()

exit(0)