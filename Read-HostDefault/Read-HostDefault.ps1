param(
 [object]$Prompt,
 [object]$Default)
 
[void][System.Windows.Forms.SendKeys]
 [System.Windows.Forms.SendKeys]::SendWait(
 ([regex]'([\{\}\[\]\(\)\+\^\%\~])').Replace($Default, '{$1}'))
 
Read-Host -Prompt $Prompt
 
trap {
 [void][System.Reflection.Assembly]::LoadWithPartialName(
 'System.Windows.Forms')
 continue
 }
