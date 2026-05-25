param(
    [string]$msg = "update"
)
git add .
git commit -m $msg
git push
