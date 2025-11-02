<#
  Set-LegacyCipherSuite.ps1
  Reinstates TLS 1.0/1.1 compatibility for legacy SQL clients while keeping TLS 1.3 disabled.
  Run from an elevated session.
#>

$policyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002'
$cipherList = @(
    'TLS_AES_256_GCM_SHA384','TLS_AES_128_GCM_SHA256',
    'TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384','TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256',
    'TLS_RSA_WITH_AES_256_GCM_SHA384','TLS_RSA_WITH_AES_128_GCM_SHA256',
    'TLS_RSA_WITH_AES_256_CBC_SHA256','TLS_RSA_WITH_AES_128_CBC_SHA256',
    'TLS_RSA_WITH_AES_256_CBC_SHA','TLS_RSA_WITH_AES_128_CBC_SHA'
)

if (-not (Test-Path $policyKey)) {
    New-Item -Path $policyKey -Force | Out-Null
}

New-ItemProperty -Path $policyKey -Name 'Functions' -PropertyType MultiString -Value $cipherList -Force | Out-Null

Write-Host "Cipher suite order applied. Reboot required for services to pick up the change." -ForegroundColor Green
