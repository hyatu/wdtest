Get-ChildItem -Directory -Path C:\Technet\Management -Recurse -Depth 5 |  
    Foreach-Object {  
        $Acl = Get-Acl -Path $_.FullName  
        foreach ($Access in $acl.Access) {  
            $Properties = [ordered]@{  
                                        'FolderName'        = $_.FullName  
                                        'AD Group or User'  = $Access.IdentityReference  
                                        'Permissions'       = $Access.FileSystemRights  
                                        'Inherited'         = $Access.IsInherited  
                                        'LastWriteTime'     = $_.LastWriteTime  
                                    }  
            [PSCustomObject]$Properties  
        }  
    } | Export-Csv -Path "C:\temp\FolderPermissions.csv"  