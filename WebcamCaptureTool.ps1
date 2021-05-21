Set-StrictMode -Version Latest

# �c�[����
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

# �f�o�b�O���[�h
$DebugPreference = "SilentlyContinue"
#$DebugPreference = "Continue"

# �X�N���v�g�t�H���_�ւ̈ړ�
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
Write-Debug "> Script Directory: $ScriptDir"
Set-Location $ScriptDir

# �摜�ۑ���
$SaveDirectory = $ScriptDir

Write-Debug "----------------------------------------------------------------------"
Write-Debug "�����"
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
Write-Debug "OpenCvSharpExtern.dll �̃R�s�["
$OpenCvSharpExternDLLFileName = "OpenCvSharpExtern.dll"
Copy-Item $(Join-Path -Path $LibDir -ChildPath "$OpenCvSharpExternDLLFileName") $(Join-Path -Path $ScriptDir -ChildPath "$OpenCvSharpExternDLLFileName") -Force

Write-Debug "----------------------------------------------------------------------"
Write-Debug "�A�Z���u���̃��[�h"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Debug "----------------------------------------------------------------------"
Write-Debug "OpenCvSharp �̃��[�h"
$OpenCvSharpDLL = Join-Path -Path $ScriptDir -ChildPath "OpenCvSharp.dll"
$Res = [System.Reflection.Assembly]::LoadFrom($OpenCvSharpDLL)
Write-Debug "> $Res"
$OpenCvSharpExtDLL = Join-Path -Path $ScriptDir -ChildPath "OpenCvSharp.Extensions.dll"
$Res = [System.Reflection.Assembly]::LoadFrom($OpenCvSharpExtDLL)
Write-Debug "> $Res"

Write-Debug "----------------------------------------------------------------------"
Write-Debug "�A�Z���u�����[�h��"
$Res = [System.AppDomain]::CurrentDomain.GetAssemblies() | % { $_.GetName().Name }
foreach ($record in $Res) {
    Write-Debug "> $record"
}

Write-Debug "----------------------------------------------------------------------"
Write-Debug "�J�����֘A�֐��̏���"

$LoadDialogWidth = 320
$LoadDialogHeight = 64
 
$global:LoadForm = New-Object System.Windows.Forms.Form
$global:LoadForm.ClientSize = New-Object System.Drawing.Size($LoadDialogWidth, $LoadDialogHeight)
$global:LoadForm.StartPosition = "CenterScreen"
$global:LoadForm.AutoSize = $False
$global:LoadForm.FormBorderStyle = "None"
$global:LoadForm.MaximizeBox = $False
$global:LoadForm.MinimizeBox = $False
$global:LoadForm.Topmost = $True

$global:LoadMessageLabel = New-Object System.Windows.Forms.Label
$global:LoadMessageLabel.Location = New-Object System.Drawing.Point(0, 0)
$global:LoadMessageLabel.Size = New-Object System.Drawing.Size($LoadDialogWidth, $LoadDialogHeight)
$global:LoadMessageLabel.Text = "�N����..."
$global:LoadMessageLabel.Font = [System.Drawing.Font]::new("Meiryo UI", 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, 128)
$global:LoadMessageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$global:LoadForm.Controls.Add($global:LoadMessageLabel)

[void] $global:LoadForm.Show()

$CaptureWidth = [int]1280
$CaptureHeight = [int]720

$global:Capture = [OpenCvSharp.VideoCapture]::new()
[void]$global:Capture.Open(0)
[void]$global:Capture.Set([OpenCvSharp.VideoCaptureProperties]::FrameWidth, $CaptureWidth)
[void]$global:Capture.Set([OpenCvSharp.VideoCaptureProperties]::FrameHeight, $CaptureHeight)

$global:LoadForm.Close()

if (-not $global:Capture.IsOpened()) {
    Write-Host "�J�����̃I�[�v���Ɏ��s���܂����B"
    FinalizeCamera
    exit(-1)
}

function FinalizeCamera() {
    Write-Debug "�J�����̏I��"
    if ($global:Capture.IsOpened()) {
        $global:Capture.Release()
        $global:Capture.Dispose()
    }
}

function Capture() {
    Write-Debug "�J��������t���[�����擾"
    $Frame = [OpenCvSharp.Mat]::new()
    $global:Capture.Read($Frame)
    $global:Label.Text = "�B�e�ɐ������܂����B�\������Ă���摜��ۑ�����ꍇ�ɂ� [�ۑ�] ���N���b�N���Ă��������B"
    if ($Frame.Empty()) {
        Write-Host "�t���[���̎擾�Ɏ��s"
        $global:Label.Text = "�G���[: �t���[���̎擾�Ɏ��s���܂����B"
        FinalizeCamera
        exit(-1)
    }
    $Bitmap = [OpenCvSharp.Extensions.BitmapConverter]::ToBitmap($Frame)
    $global:PictureBox.Image = $Bitmap
}

function SaveAsFile() {
    Write-Debug "�t�@�C���ŕۑ�"
    $Bitmap = $global:PictureBox.Image

    $FilePath = Join-Path -Path $SaveDirectory -ChildPath $(GenerateFilename)
    Write-Debug "> $FilePath"

    $Bitmap.Save($FilePath)

    if (Test-Path $FilePath) {
        $global:Label.Text = "�ۑ��ɐ������܂����B (�ۑ���: $FilePath)"
    } else {
        $global:Label.Text = "�G���[: �ۑ��Ɏ��s���܂����B (�ۑ���: $FilePath)"
        [System.Windows.Forms.MessageBox]::Show("�ۑ��Ɏ��s���܂����B�B�e���������Ă��邩�A�ۑ���ɃA�N�Z�X�ł��邩�m�F���Ă��������B", "�G���[") 
    }
}

function GenerateFilename() {
    $Filename = "Capture" + "_" + $env:USERNAME + "_" + $env:COMPUTERNAME+ "_" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".png"
    return [string]$Filename
}

Write-Debug "----------------------------------------------------------------------"
Write-Debug "�E�C���h�E�̐���"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles()

$Margin = [int]24
$ButtonWidth = [int]150
$ButtonHeight = [int]32
$LabelWidth = [int]($CaptureWidth - $Margin * 2)
$LabelHeight = [int]24

# �t�H�[��
$global:MainForm = New-Object System.Windows.Forms.Form
$global:MainForm.ClientSize = New-Object System.Drawing.Size([int]($CaptureWidth + $Margin * 2), [int]($CaptureHeight + $ButtonHeight + $Margin * 3))
$global:MainForm.StartPosition = "CenterScreen"
$global:MainForm.AutoSize = $False
$global:MainForm.FormBorderStyle = "FixedSingle"
$global:MainForm.MaximizeBox = $False
$global:MainForm.MinimizeBox = $False
$global:MainForm.Text = "$AppName ($env:USERNAME@$env:COMPUTERNAME)"

# ���j���[�X�g���b�v
$MenuStrip = New-Object System.Windows.Forms.MenuStrip
$global:MainForm.MainMenuStrip = $MenuStrip
$global:MainForm.Controls.Add($MenuStrip)

# ���j���[ - �B�e
$CaptureMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$CaptureMenuItem.Text = "�B�e(&W)"
$CaptureMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::W
$CaptureMenuItem.ShowShortcutKeys = $True
$CaptureMenuItem.Add_Click({
    Capture
})
[void]$MenuStrip.Items.Add($CaptureMenuItem)

# ���j���[ - �ۑ�
$SaveAsFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$SaveAsFileMenuItem.Text = "�ۑ�(&S)"
$SaveAsFileMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::S
$SaveAsFileMenuItem.Add_Click({
    SaveAsFile
})
[void]$MenuStrip.Items.Add($SaveAsFileMenuItem)

# ���j���[ - �I��
$QuitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$QuitMenuItem.Text = "�I��(&Q)"
$QuitMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::Q
$QuitMenuItem.Add_Click({
    $global:MainForm.Close()
})
[void]$MenuStrip.Items.Add($QuitMenuItem)

# ���j���[ - �A�v���ɂ���
$AboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutMenuItem.Text = "�A�v���ɂ���(&A)"
$AboutMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::A
$AboutMenuItem.Add_Click({ 
    
    $AboutFormMargin = [int]24
    $AboutFormWidth = [int]320
    $AboutFormHeight = [int]240
    $AboutLabelHeight = [int]36

    $AboutForm = New-Object System.Windows.Forms.Form
    $AboutForm.ClientSize = New-Object System.Drawing.Size($AboutFormWidth, $AboutFormHeight)
    $AboutForm.StartPosition = "CenterScreen"
    $AboutForm.AutoSize = $False
    $AboutForm.FormBorderStyle = "FixedSingle"
    $AboutForm.MaximizeBox = $False
    $AboutForm.MinimizeBox = $False
    $AboutForm.Text = "�A�v���ɂ���"
    $AboutForm.Owner = $global:MainForm

    # �A�v����
    $AboutAppNameLabel = New-Object System.Windows.Forms.Label
    $AboutAppNameLabel.Location = New-Object System.Drawing.Point([int]($AboutFormMargin), [int]($AboutFormMargin))
    $AboutAppNameLabel.Size = New-Object System.Drawing.Size([int]($AboutFormWidth - $AboutFormMargin * 2), [int]($AboutLabelHeight))
    $AboutAppNameLabel.Text = $AppName
    $AboutAppNameLabel.Font = [System.Drawing.Font]::new("MS UI Gothic", 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, 128)
    $AboutAppNameLabel.TextAlign = "MiddleCenter"
    $AboutForm.Controls.Add($AboutAppNameLabel)

    # �R�s�[���C�g
    $AboutCopyrightLabel = New-Object System.Windows.Forms.Label
    $AboutCopyrightLabel.Location = New-Object System.Drawing.Point([int]($AboutFormMargin), [int]($AboutFormMargin + $AboutLabelHeight))
    $AboutCopyrightLabel.Size = New-Object System.Drawing.Size([int]($AboutFormWidth - $AboutFormMargin * 2), [int]($AboutLabelHeight))
    $AboutCopyrightLabel.Text = "Copyright (c) 2021 tag. All rights reserved."
    $AboutCopyrightLabel.Font = [System.Drawing.Font]::new("MS UI Gothic", 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 128)
    $AboutCopyrightLabel.TextAlign = "MiddleCenter"
    $AboutForm.Controls.Add($AboutCopyrightLabel)

    # ���ϐ�
    $AboutCopyrightLabel = New-Object System.Windows.Forms.Label
    $AboutCopyrightLabel.Location = New-Object System.Drawing.Point([int]($AboutFormMargin), [int]($AboutFormMargin + $AboutLabelHeight * 2))
    $AboutCopyrightLabel.Size = New-Object System.Drawing.Size([int]($AboutFormWidth - $AboutFormMargin * 2), [int]($AboutLabelHeight * 3))
    $AboutCopyrightLabel.Text = "Powered by OpenCvSharp`n`nPowerShell $Version ($Architecture)"
    $AboutCopyrightLabel.Font = [System.Drawing.Font]::new("MS UI Gothic", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 128)
    $AboutCopyrightLabel.TextAlign = "MiddleCenter"
    $AboutForm.Controls.Add($AboutCopyrightLabel)

    [void] $AboutForm.Show()
})
[void]$MenuStrip.Items.Add($AboutMenuItem)

# �s�N�`���[�{�b�N�X
$global:PictureBox = New-Object System.Windows.Forms.PictureBox
$global:PictureBox.Location = New-Object System.Drawing.Point($Margin, $Margin)
$global:PictureBox.Size = New-Object System.Drawing.Size($CaptureWidth, $CaptureHeight)
$global:MainForm.Controls.Add($PictureBox)

# ���x����\��
$global:Label = New-Object System.Windows.Forms.Label
$global:Label.Location = New-Object System.Drawing.Point([int]($Margin), [int]($CaptureHeight + $Margin * 2))
$global:Label.Size = New-Object System.Drawing.Size($LabelWidth, $LabelHeight)
$global:Label.Text = "[�B�e] ���N���b�N���Ă��������B"
$global:Label.TextAlign = "MiddleLeft"
$global:MainForm.Controls.Add($global:Label)

# �t�H�[���̃A�N�e�B�x�[�g
$global:MainForm.Add_Shown({
    #Capture
    $global:MainForm.Activate()
})
[void] $global:MainForm.ShowDialog()

FinalizeCamera
exit(0)