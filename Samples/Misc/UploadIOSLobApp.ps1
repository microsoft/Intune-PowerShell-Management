# Create the object that contains information about the app.

$appToUpload = New-MobileAppObject `
    -iosLobApp `
    -displayName 'Test' `
    -description 'This is a test iOS LOB app' `
    -publisher 'Rohit Ramu' `
    -bundleId '' `
    -applicableDeviceType (New-IosDeviceTypeObject -iPad $true -iPhoneAndIPod $true) `
    -minimumSupportedOperatingSystem (New-IosMinimumOperatingSystemObject -v9_0 $true) `
    -fileName 'test.ipa' `
    -buildNumber 'v1' -versionNumber 'v1' -expirationDateTime ((Get-Date).AddDays(90))

# Upload the app file with the app information.
# !! Replace 'test.ipa' with the path to your *.ipa file !!

$createdApp = New-LobApp `
    -filePath 'test.ipa' `
    -mobileApp $appToUpload

Write-Output $createdApp
