param (
    [Parameter(Mandatory=$true)]
    [string]$NewVersion
)

# version.json 업데이트
$versionFile = "version.json"
$version = Get-Content $versionFile | ConvertFrom-Json
$version.version = $NewVersion
$version.releaseDate = (Get-Date).ToString("yyyy-MM-dd")
$version | ConvertTo-Json | Set-Content $versionFile

# RELEASE_NOTES.md 업데이트
$releaseNotes = Get-Content "RELEASE_NOTES.md"
$releaseNotes[0] = "# ScreensaverAuditor v$NewVersion"
$releaseNotes | Set-Content "RELEASE_NOTES.md"

# git 태그 생성 및 푸시
git add version.json RELEASE_NOTES.md
git commit -m "Update version to v$NewVersion"
git tag "v$NewVersion"
git push
git push --tags
