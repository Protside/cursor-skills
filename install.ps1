[CmdletBinding(SupportsShouldProcess = $true)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-ProtectedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RelativePath
    )

    $parts = $RelativePath -split '[\\/]'
    if ($parts | Where-Object { $_ -in @('.git', 'secrets', 'credentials') }) {
        return $true
    }

    $name = $parts[-1]
    return (
        $name -eq '.env' -or
        $name -like '.env.*' -or
        $name -ieq 'oauth-client.json' -or
        $name -like 'token-*.json' -or
        $name -ieq 'credentials.json' -or
        $name -like 'client_secret*.json'
    )
}

try {
    $repositoryRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $skillsRoot = Join-Path $repositoryRoot 'skills'
    $globalSkillsRoot = Join-Path $env:USERPROFILE '.cursor\skills'

    if (-not (Test-Path -LiteralPath $skillsRoot -PathType Container)) {
        throw "Repository skills directory does not exist: $skillsRoot"
    }

    $skillDirectories = @(
        Get-ChildItem -LiteralPath $skillsRoot -Directory -Force |
            Sort-Object -Property Name
    )

    if ($skillDirectories.Count -eq 0) {
        throw "No skill directories were found under: $skillsRoot"
    }

    if (-not (Test-Path -LiteralPath $globalSkillsRoot -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($globalSkillsRoot, 'Create global Cursor skills directory')) {
            New-Item -ItemType Directory -Path $globalSkillsRoot -Force | Out-Null
        }
    }

    foreach ($skill in $skillDirectories) {
        $destinationRoot = Join-Path $globalSkillsRoot $skill.Name
        $sourceItems = @(Get-ChildItem -LiteralPath $skill.FullName -Recurse -Force)
        $protectedItems = @()

        foreach ($item in $sourceItems) {
            $relativePath = $item.FullName.Substring($skill.FullName.Length).TrimStart('\', '/')
            if (Test-ProtectedPath -RelativePath $relativePath) {
                $protectedItems += $relativePath
                continue
            }

            $destinationPath = Join-Path $destinationRoot $relativePath

            if ($item.PSIsContainer) {
                if (
                    -not (Test-Path -LiteralPath $destinationPath -PathType Container) -and
                    $PSCmdlet.ShouldProcess($destinationPath, 'Create skill directory')
                ) {
                    New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
                }
                continue
            }

            $destinationDirectory = Split-Path -Parent $destinationPath
            if (
                -not (Test-Path -LiteralPath $destinationDirectory -PathType Container) -and
                $PSCmdlet.ShouldProcess($destinationDirectory, 'Create skill directory')
            ) {
                New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
            }

            if ($PSCmdlet.ShouldProcess($destinationPath, "Install skill file from $($item.FullName)")) {
                $temporaryPath = Join-Path $destinationDirectory (
                    ".{0}.cursor-install-{1}.tmp" -f $item.Name, [guid]::NewGuid().ToString('N')
                )

                try {
                    Copy-Item -LiteralPath $item.FullName -Destination $temporaryPath -Force
                    Move-Item -LiteralPath $temporaryPath -Destination $destinationPath -Force
                }
                finally {
                    if (Test-Path -LiteralPath $temporaryPath) {
                        Remove-Item -LiteralPath $temporaryPath -Force
                    }
                }
            }
        }

        if ($protectedItems.Count -gt 0) {
            Write-Warning (
                "Skipped protected paths in skill '{0}': {1}" -f
                $skill.Name,
                ($protectedItems -join ', ')
            )
        }

        if ($WhatIfPreference) {
            Write-Host "Validated skill (WhatIf): $($skill.Name)"
        }
        else {
            Write-Host "Installed skill: $($skill.Name)"
        }
    }
}
catch {
    Write-Error "Skill installation failed: $($_.Exception.Message)"
    exit 1
}
