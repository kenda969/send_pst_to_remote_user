Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 
dir “C:\  \.pst” | % { $namespace.AddStore($_.FullName) }
write-host $namespace
