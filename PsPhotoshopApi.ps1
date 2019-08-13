Function Unregister-ComObject($Reference, $ObjectName) {
try { 
    if ($Reference) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Reference) | out-null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 500
        $Reference = $null
    } else { Write-Host "WARNING! Object '$Objectname' was not instantiated before.! $_" -foregroundcolor "magenta" }
    } catch { 
        Write-Host "ERROR! Exception occured, releasing object '$Objectname' failed! $_" -foregroundcolor "red" 
    }
}


$WinApi = @"
using System;
using System.Runtime.InteropServices;
public class WinApi {
    [DllImport("user32.dll")]
    public static extern int PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
}
"@
Add-Type $WinApi


function Set-CloseWindow($proc) { 
    try {
        $p = Get-Process | Where-Object {$_.mainWindowTItle } | Where-Object {$_.Name -like "$proc"}
        if ($p) {
            $h = $p.MainWindowHandle
            if ($h) {
            [void] [WinApi]::PostMessage($h, 0x10, 0, 0)
            } else { Write-Host "WARNING: No valid handle found for '$Proc'" -foregroundcolor "magenta" }
        } else { Write-Host "WARNING: No process found for '$Proc'" -foregroundcolor "magenta" }
     } catch {
        Write-Host "ERROR: Exception occured, the application '$Proc' could not be closed. $_" -foregroundcolor "red"
    }
}


Function Open-Photoshop() {
    [int] $PsPixels = 1
    [int] $PsDisplayNoDialogs = 3

    try {
        $Script:PsObj = New-Object -ComObject "Photoshop.Application.130"
        $Script:ExportObj = New-Object -ComObject "Photoshop.ExportOptionsSaveForWeb.130"

        # Save Defaults
        $Script:OldRulerUnits = $Script:PsObj.Preferences.RulerUnits 
        $Script:OldTypeUnits = $Script:PsObj.Preferences.TypeUnits 
        $Script:OldDisplayDialogs = $Script:PsObj.DisplayDialogs
        
        # Set our Defaults
        $Script:PsObj.Preferences.RulerUnits = $PsPixels
        $Script:PsObj.Preferences.TypeUnits = $PsPixels
        $Script:PsObj.DisplayDialogs = $PsDisplayNoDialogs 
        Return $True
    } catch {
        Write-Host "ERROR: Exception occured creating photoshop.application object. $_" -foregroundcolor "red"
        Return $false
    }
}


Function Close-Photoshop() {
    # Restore Defaults
    if ($Script:PsObj) {
        $Script:PsObj.Preferences.RulerUnits = $Script:OldRulerUnits
        $Script:PsObj.Preferences.TypeUnits = $Script:OldTypeUnits
        $Script:PsObj.DisplayDialogs = $Script:OldDisplayDialogs 
    }    

    Unregister-ComObject $Script:ExportObj "Photoshop.ExportOptionsSaveForWeb"
    Unregister-ComObject $Script:PsObj "Photoshop.Application"
    Set-CloseWindow "photoshop" 
}


Function Export-Png($FilePath) {
    try {
        [int] $PsSavePNG = 13
        [int] $PsExportSAVEFORWEB = 2
        [int] $PsDoNotSaveChanges = 2
        $Doc = $Script:PsObj.Open($FilePath)
        $Script:ExportObj.format = $PsSavePNG
        $Script:ExportObj.PNG8 = $True
        $Script:ExportObj.Transparency = $True
        $Script:ExportObj.Interlaced = $True
        $Script:ExportObj.Quality = 8
        $FName = [io.path]::ChangeExtension($FilePath, "png") 
        $Doc.Export($FName,$PsExportSAVEFORWEB, $Script:ExportObj)
        $Script:PsObj.ActiveDocument.Close($PsDoNotSaveChanges)
    } catch {
        Write-Host "ERROR: Exception occured on exporting png for the web. $_" -foregroundcolor "red"
        Return $false 
    } 
    Return $true
}


if (Open-Photoshop) {
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    Export-Png "$DesktopPath\test.psd" | Out-Null
    Close-Photoshop
}