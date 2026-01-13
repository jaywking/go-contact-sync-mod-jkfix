param(
    [switch]$release
)

# FUNCTIONS

function CenterText {
    param(
        $Message, 
        $width
    )
    $string = "= "
    for($i = 0; $i -lt (([Math]::Max(0, $width / 2) - [Math]::Max(0, $Message.Length / 2))) - 4; $i++)
    {
        $string = $string + " "
    }
    $string = $string + $Message
    for($i = 0; $i -lt ($width - ((([Math]::Max(0, $width / 2) - [Math]::Max(0, $Message.Length / 2))) - 2 + $Message.Length)) - 2; $i++)
    {
        $string = $string + " "
    }
    $string = $string + " ="
    return $string
}

function LeftText {
    param(
        $Message,
        $width
    )
    return (("= $Message".PadRight($width-1," "))+"=")
}

function WriteLine {
    param(
        $width
    )
    return (("=".PadRight($width-1,"="))+"=")
}

# VARIABLES

$sfhost = "frs.sourceforge.net"
$baseDirectory = "/home/frs/project/googlesyncmod/"
$width = 80 #$Host.UI.RawUI.BufferSize.Width

$sfuser = Read-Host "Sourceforge username"
$secureSfPass = Read-Host "Sourceforge password" -AsSecureString

WriteLine $width
if($release) {
    $uploadDirectory = $baseDirectory + "Releases/"
    $versionXmlDirectory = $baseDirectory

    CenterText "RELEASE MODE" $width
}
else {
    $uploadDirectory = $baseDirectory + "Utilities/test/"
    $versionXmlDirectory = $uploadDirectory

    CenterText "TEST MODE" $width
}

WriteLine $width
LeftText ("HOST: " + $sfhost) $width
LeftText ("USER: " + $sfuser) $width
LeftText ("VER:  <release>") $width
LeftText ("DIR:  " + $uploadDirectory) $width
LeftText ("XML:  " + $versionXmlDirectory) $width
WriteLine $width

Set-Location -Path $PSScriptRoot

$msg = "Upload files? [y/n]"

while ($True) {
    $response = Read-Host -Prompt $msg
    if ($response -match "[yY]") {
        Write-Output "`nCopying directory <release> to $uploadDirectory ..."
        Start-Process -FilePath "../../Tools/pscp.exe" -ArgumentList "-r -p -scp -pw $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSfPass))) <release> $sfuser@${sfhost}:$uploadDirectory" -wait -NoNewWindow
    
        Write-Output "`nCopying version xml for auto-updater to $versionXmlDirectory ..."
        Start-Process -FilePath "../../Tools/pscp.exe" -ArgumentList "-p -scp -pw $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSfPass))) updates_v1.xml $sfuser@${sfhost}:$versionXmlDirectory" -wait -NoNewWindow
        
        #Write-Output "`nCopying version information README.md ..."
        #Start-Process -FilePath "../../Tools/pscp.exe" -ArgumentList "-p -scp -pw $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSfPass))) README.md $sfuser@$host:$uploadDirectory" -wait -NoNewWindow
        break
    } elseif ($response -match "[nN]") {
        Write-Host "Upload canceled."
        break
    } else {
        continue
    }
}
