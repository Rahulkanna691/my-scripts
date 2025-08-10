# Load Excel module
Import-Module ImportExcel -ErrorAction Stop

# Excel file path
$excelPath = "C:\Users\List of Sarboxprojects.xlsx"

# PAT Token
$pat = #"PAT"
$base64AuthInfo = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$pat"))
Add-Type -AssemblyName System.Web

# Permission Map
$permissionMap = @{
    1     = "Read"
    2     = "Contribute"
    4     = "Administer"
    8     = "ForcePush"
    16    = "CreateBranch"
    32    = "CreateTag"
    64    = "ManageNote"
    128   = "PolicyExempt"
    256   = "PullRequestContribute"
    512   = "CreateRepository"
    1024  = "DeleteRepository"
}

function Resolve-Bitmask {
    param ([int]$bitmask)
    return ($permissionMap.Keys | Where-Object { $bitmask -band $_ } | ForEach-Object { $permissionMap[$_] }) -join ", "
}

function Get-DisplayNameFromDescriptor {
    param (
        [string]$rawDescriptor,
        [string]$organization,
        [hashtable]$headers,
        [hashtable]$displayNameCache
    )

    if ($displayNameCache.ContainsKey($rawDescriptor)) {
        return $displayNameCache[$rawDescriptor]
    }

    if ($rawDescriptor -match "Microsoft.IdentityModel.Claims.ClaimsIdentity;(.+\\)?(.+@.+)$") {
        $upn = $matches[2]
        $displayNameCache[$rawDescriptor] = $upn
        return $upn
    }

    $encodedDescriptor = [System.Web.HttpUtility]::UrlEncode($rawDescriptor)
    $subjectUrl = "https://vssps.dev.azure.com/$organization/_apis/graph/subjectquery?api-version=7.0-preview.1"
    $body = @{ subjects = @($encodedDescriptor) } | ConvertTo-Json -Depth 2

    try {
        $resp = Invoke-RestMethod -Uri $subjectUrl -Headers $headers -Method Post -Body $body -ContentType "application/json"
        $name = $resp.subjects[0].displayName
        if (-not $name) { $name = $rawDescriptor }
    } catch {
        $name = $rawDescriptor
    }

    $displayNameCache[$rawDescriptor] = $name
    return $name
}

# Read Excel rows
$projectList = Import-Excel -Path $excelPath

# Loop through each org/project
foreach ($row in $projectList) {
    $organization = $row.Organization
    $project = $row.Project
    $headers = @{
        Authorization = "Basic $base64AuthInfo"
        Accept        = "application/json"
    }
    $displayNameCache = @{}
    $csvOutput = @()

    Write-Host "`n📌 Processing Organization: $organization | Project: $project"

    try {
        # Get Project ID
        $projectUrl = "https://dev.azure.com/$organization/_apis/projects?api-version=7.0"
        $projectObj = (Invoke-RestMethod -Uri $projectUrl -Headers $headers).value | Where-Object { $_.name -eq $project }
        if (-not $projectObj) {
            Write-Warning "⚠️ Project '$project' not found in $organization!"
            continue
        }
        $projectId = $projectObj.id 

        # Get Repositories
        $repoListUrl = "https://dev.azure.com/$organization/$project/_apis/git/repositories?api-version=7.0"
        $repos = Invoke-RestMethod -Uri $repoListUrl -Headers $headers

        $securityNamespaceId = "2e9eb7ed-3c0a-47d4-87c1-0ffdd275fd87"

        foreach ($repo in $repos.value) {
            $repoId = $repo.id
            $repoName = $repo.name
            $securityToken = "repoV2/$projectId/$repoId"

            Write-Host "`n============================="
            Write-Host "Repo Name : $repoName"
            Write-Host "Repo ID   : $repoId"
            Write-Host "Token     : $securityToken"
            Write-Host "============================="

            $aclUrl = "https://dev.azure.com/$organization/_apis/accesscontrollists/2e9eb7ed-3c0a-47d4-87c1-0ffdd275fd87?token=repoV2/$($projectId)/$($repoId)&includeExtendedInfo=true&recurse=true"

            try {
                $aclResponse = Invoke-RestMethod -Uri $aclUrl -Headers $headers -Method Get

                if ($aclResponse.value -and $aclResponse.value[0].acesDictionary) {
                    Write-Host "`nPermissions Found:`n"
                    $repoPermissions = @()

                    foreach ($entry in $aclResponse.value[0].acesDictionary.PSObject.Properties) {
                        $descriptor = $entry.Name
                        $ace = $entry.Value

                        if ($descriptor -like "Microsoft.TeamFoundation.ServiceIdentity;*" -or $descriptor -like "Microsoft.TeamFoundation.Identity;*") {
                            continue
                        }

                        $displayName = Get-DisplayNameFromDescriptor -rawDescriptor $descriptor -organization $organization -headers $headers -displayNameCache $displayNameCache
                        $allow = Resolve-Bitmask -bitmask $ace.allow
                        $deny  = Resolve-Bitmask -bitmask $ace.deny

                        $repoPermissions += [PSCustomObject]@{
                            'Display Name' = $displayName
                            'Allow Rights' = $allow
                            'Deny Rights'  = $deny
                        }

                        $csvOutput += [PSCustomObject]@{
                            'Organization' = $organization
                            'Project'      = $project
                            'Repo Name'    = $repoName
                            'Repo ID'      = $repoId
                            'Token'        = $securityToken
                            'Display Name' = $displayName
                            'Allow Rights' = $allow
                            'Deny Rights'  = $deny
                        }
                    }

                    if ($repoPermissions.Count -gt 0) {
                        $repoPermissions | Sort-Object 'Display Name' | Format-Table -AutoSize
                    } else {
                        Write-Host "No explicit permissions to display for this repo.`n"
                    }
                } else {
                    Write-Host "⚠️ No ACL entries for this repository."
                }

            } catch {
                Write-Warning "Failed to fetch permissions for repo '$repoName': $_"
            }
        }

        # Export to CSV (append for each project)
        $date = Get-Date -Format 'yyyyMMdd'
        $outputPath = "C:\Users\AzureRepoPermissions_$date.csv"
        $csvOutput | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Append

        Write-Host "`n✅ Exported results for $organization / $project"

    } catch {
        Write-Error "💥 Failed on Organization: $organization | Project: $project -> $_"
    }
}
