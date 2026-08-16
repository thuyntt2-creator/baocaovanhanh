$sh = New-Object -ComObject Shell.Application
$rb = $sh.Namespace(10)

Write-Host "Searching Recycle Bin..."
$found = $false
foreach ($item in $rb.Items()) {
    if ($item.Name -like "*New folder*" -or $item.Name -like "*dashboard*") {
        Write-Host "FOUND DELETED ITEM:" $item.Name "at" $item.Path
        $found = $true
        # Attempt restore (InvokeVerb Undelete / restore)
        try {
            $item.InvokeVerb("restore")
            Write-Host "✅ SUCCESSFULLY RESTORED:" $item.Name
        } catch {
            Write-Host "❌ Could not auto-restore verb, item path:" $item.Path
        }
    }
}

if (-not $found) {
    Write-Host "Recycle bin items:"
    foreach ($item in $rb.Items()) {
        Write-Host " - " $item.Name
    }
}
