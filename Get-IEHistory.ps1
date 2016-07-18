function get-iehistory {            
  [CmdletBinding()]            
  param ()            
            
  $shell = New-Object -ComObject Shell.Application            
  $hist = $shell.NameSpace(34)            
  $folder = $hist.Self            
            
  $hist.Items() |             
  foreach {            
    if ($_.IsFolder) {            
      $siteFolder = $_.GetFolder            
      $siteFolder.Items() |             
      foreach {            
        $site = $_            
             
        if ($site.IsFolder) {            
          $pageFolder  = $site.GetFolder            
          $pageFolder.Items() |             
          foreach {            
            $visit = New-Object -TypeName PSObject -Property @{            
              Site = $($site.Name)            
              URL = $($pageFolder.GetDetailsOf($_,0))            
              Date = $( $pageFolder.GetDetailsOf($_,2))            
            }            
            $visit            
          }            
        }            
      }            
    }            
  }            
}
