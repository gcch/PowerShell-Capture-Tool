# PowerShell Webcam Capture Tool

## Overview

PowerShell + OpenCvSharp にて PC に接続された Web カメラで画像撮影し、ファイル保存するサンプルスクリプト

## Screenshots

<div align="center">
  <img src="https://github.com/gcch/PowerShell-Webcam-Capture-Tool/blob/main/Screenshots/PowerShell-Webcam-Capture-Tool_ss-01.jpg" alt="" title="PowerShell-Webcam-Capture-Tool_ss-01">
</div>

## Environment

スクリプト作成環境は以下の通り

- Windows 10 Version 20H2
- PowerShell 5.1.19041.906 (64-bit)
- OpenCvSharp 4.5.2 (64-bit)

## Setup
以下の手順で準備してください。

1. 本スクリプトをダウンロードし、任意のフォルダ (以降、&lt;スクリプトフォルダ&gt; と記載) に解凍する
1. [OpenCvSharp](https://github.com/shimat/opencvsharp/releases) からバイナリデータ (OpenCvSharp-x.y.z-YYYYMMDD.zip) をダウンロード
1. ダウンロードした "OpenCvSharp-x.y.z-YYYYMMDD.zip" を解凍 (以降、&lt;OpenCvSharp 解凍フォルダ&gt; と記載)
1. "&lt;OpenCvSharp 解凍フォルダ&gt;\ManagedLib\OpenCvSharp.dll" 及び 、"&lt;OpenCvSharp 解凍フォルダ&gt;\ManagedLib\OpenCvSharp.Extensions.dll" を &lt;スクリプトフォルダ&gt; 直下にコピー
1. "&lt;OpenCvSharp 解凍フォルダ&gt;\NativeLib\win\x86\OpenCvSharpExtern.dll" を "&lt;スクリプトフォルダ&gt;\x86" にコピー
1. "&lt;OpenCvSharp 解凍フォルダ&gt;\NativeLib\win\x64\OpenCvSharpExtern.dll" を "&lt;スクリプトフォルダ&gt;\x64" にコピー

## Setting
「WebcamCaptureTool.ps1」を編集し、画像保存先フォルダを設定してください。デフォルトではスクリプトと同じフォルダが設定されています。

```
$SaveDirectory = <保存先フォルダパス>
```

## Launch
「WebcamCaptureTool.vbs」をダブルクリップして起動してください。

## References
- [GitHub - shimat/opencvsharp: OpenCV wrapper for .NET](https://github.com/shimat/opencvsharp)

## License
All components are licensed under Apache License 2.0.

## Author
tag (Twitter: [@tag_ism](https://twitter.com/tag_ism "tag (@tag_ism) | Twitter") / Blog: http://karat5i.blogspot.jp/)

