# Path to the central wallpaper share
$SharePath = "\\DC02\DistroFolder\Wallpapers"

# Pick a random image file from the share
$Wallpaper = Get-ChildItem -Path $SharePath -Include *.jpg, *.png -Recurse | Get-Random

if ($Wallpaper) {
    # Set the wallpaper registry key
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'Wallpaper' -Value $Wallpaper.FullName
    
    # Refresh system parameters to apply immediately without relogging
    $code = @"
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
"@
    $type = Add-Type -MemberDefinition $code -Name "Win32SystemParametersInfo" -Namespace "Win32Functions" -PassThru
    $type::SystemParametersInfo(0x0014, 0, $Wallpaper.FullName, 0x0001 -bor 0x0002)
}