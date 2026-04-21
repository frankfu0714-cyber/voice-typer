# Push VoiceTyper to the github.com/frankfu0714-cyber/voice-typer repo.
#
# Run once from inside the Onetranscribt folder:
#     .\push_to_github.ps1
#
# Uses whatever GitHub auth git is already configured with on this machine
# (the same cached creds you used for the fire-website push). If prompted,
# sign in via the Git Credential Manager window that pops up.

$ErrorActionPreference = "Stop"

$repoUrl = "https://github.com/frankfu0714-cyber/voice-typer.git"
$commitMsg = "VoiceTyper: live dictation mode, floating overlay, and accuracy tuning"

# 1. Initialise the repo if not already one
if (-not (Test-Path ".git")) {
    Write-Host "Initialising git repo..."
    git init
    git branch -M main
}

# 2. If there are already tracked files, refresh them against the current
#    .gitignore. On a fresh init there's nothing cached, so skip quietly.
$prevPref = $ErrorActionPreference
$ErrorActionPreference = "Continue"
try {
    $null = git ls-files 2>$null
    if ($LASTEXITCODE -eq 0 -and (git ls-files | Select-Object -First 1)) {
        git rm -r --cached . 2>$null | Out-Null
    }
} catch {
    # Ignore — we're about to stage everything fresh anyway.
}
$ErrorActionPreference = $prevPref

# 3. Stage everything that survives .gitignore
Write-Host "Staging files..."
git add .

# 4. Commit (skip if there's nothing to commit)
$status = git status --porcelain
if ($status) {
    Write-Host "Committing..."
    git commit -m $commitMsg
} else {
    Write-Host "Nothing to commit — working tree is clean."
}

# 5. Point origin at the GitHub repo (add if missing, otherwise update)
$existing = git remote 2>$null
if ($existing -contains "origin") {
    git remote set-url origin $repoUrl
} else {
    git remote add origin $repoUrl
}

# 6. Push
Write-Host "Pushing to $repoUrl ..."
git push -u origin main

Write-Host ""
Write-Host "Done! Check https://github.com/frankfu0714-cyber/voice-typer"
