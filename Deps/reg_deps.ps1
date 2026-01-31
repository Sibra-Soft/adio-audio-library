param(
    [Parameter(Mandatory = $true)]
    [string]$VbpFile
)

function Register($file)
{
	$Regsvr32_32 = Join-Path $env:WINDIR "SysWOW64\regsvr32.exe"
	$Regsvr32_64 = Join-Path $env:WINDIR "System32\regsvr32.exe"

    $useRegsvr32 = $Regsvr32_64
    if ($filePath -match '\\SysWOW64\\') {
        $useRegsvr32 = $Regsvr32_32
    }
	
	Write-Host "Dependency: $file"
    $p = Start-Process -FilePath $useRegsvr32 -ArgumentList "/s", "`"$file`"" -Wait -PassThru

    if ($p.ExitCode -ne 0) {
        Write-Host "  [ERROR] regsvr32 exitcode = $($p.ExitCode)" -ForegroundColor Red
    }
    else {
        Write-Host "  [OK]" -ForegroundColor Green
    }
}

if (!(Test-Path $VbpFile)) {
    Write-Host "FOUT: VBP bestand niet gevonden: $VbpFile" -ForegroundColor Red
    exit 1
}

Write-Host "======================================" 
Write-Host "VB6 Dependency Register"
Write-Host "Project: $VbpFile"
Write-Host "======================================"
Write-Host ""

$errors = 0

Register("VbLiteUnit.dll")
Register("midifl2k.ocx")
Register("midifl32.ocx")
Register("midiio2k.ocx")
Register("midiio32.ocx")

Write-Host ""
Write-Host "======================================"