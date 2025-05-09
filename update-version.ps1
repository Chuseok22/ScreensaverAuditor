param (
    [Parameter(Mandatory=$true)]
    [string]$NewVersion
)

# 시맨틱 버전 형식 검증 (예: 1.0.0)
if ($NewVersion -notmatch '^\d+\.\d+\.\d+$') {
    Write-Host "버전 형식이 올바르지 않습니다. X.Y.Z 형식이어야 합니다. (e.g., 1.0.0)."
    exit 1
}

# version.json 업데이트
$versionFile = "version.json"

# version.json 파일이 존재하는지 확인
if (-not (Test-Path $versionFile)) {
    Write-Host "version.json 파일이 존재하지 않습니다."
    exit 1
}
$version = Get-Content $versionFile | ConvertFrom-Json
$version.version = $NewVersion
$version.releaseDate = (Get-Date).ToString("yyyy-MM-dd")
$version | ConvertTo-Json | Set-Content $versionFile

# RELEASE_NOTES.md 업데이트
$releaseNotesFile = "RELEASE_NOTES.md"
# RELEASE_NOTES.md 파일이 존재하는지 확인
if (-not (Test-Path $releaseNotesFile)) {
    Write-Host "RELEASE_NOTES.md 파일이 존재하지 않습니다. 파일을 새로 생성합니다."
    "# Screensaver v$NewVersion" | Set-Content $releaseNotesFile
} else {
    $releaseNotes = Get-Content $releaseNotesFile
    $releaseNotes[0] = "# Screensaver v$NewVersion"
    $releaseNotes | Set-Content $releaseNotesFile
}

# git 태그 생성 및 푸시
function Invoke-GitCommand {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Command,
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorMessage
    )
    
    try {
        Invoke-Expression "git $Command"
        if ($LASTEXITCODE -ne 0) {
            throw "Git 명령이 코드 $LASTEXITCODE로 실패했습니다."
        }
    } catch {
        Write-Error "$ErrorMessage`n$_"
        exit 1
    }
}
    
nvoke-GitCommand -Command "add version.json RELEASE_NOTES.md" -ErrorMessage "파일 스테이징에 실패했습니다."
Invoke-GitCommand -Command "commit -m 'Update version to v$NewVersion'" -ErrorMessage "커밋 생성에 실패했습니다."
Invoke-GitCommand -Command "tag 'v$NewVersion'" -ErrorMessage "태그 생성에 실패했습니다."
Invoke-GitCommand -Command "push" -ErrorMessage "변경사항 푸시에 실패했습니다."
Invoke-GitCommand -Command "push --tags" -ErrorMessage "태그 푸시에 실패했습니다."

Write-Host "버전이 성공적으로 v$NewVersion로 업데이트되었습니다." -ForegroundColor Green
