# Self-signed HTTPS certificate generator for Interrogation App
# Run this in the project directory as Administrator

$certFile = Join-Path $PSScriptRoot "cert.pem"
$keyFile = Join-Path $PSScriptRoot "key.pem"

# Check if certs already exist
if ((Test-Path $certFile) -and (Test-Path $keyFile)) {
    Write-Host "证书已存在。如需重新生成，请先删除 cert.pem 和 key.pem 文件。" -ForegroundColor Yellow
    exit
}

# Method 1: Use OpenSSL if available
$openssl = Get-Command "openssl" -ErrorAction SilentlyContinue
if ($openssl) {
    Write-Host "使用 OpenSSL 生成证书..." -ForegroundColor Green
    & openssl req -x509 -newkey rsa:4096 -keyout $keyFile -out $certFile -days 365 -nodes -subj "/CN=localhost" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "证书生成成功！" -ForegroundColor Green
        Write-Host "  cert.pem (证书文件)" -ForegroundColor Gray
        Write-Host "  key.pem  (私钥文件)" -ForegroundColor Gray
        Write-Host "重启 start.bat 后，系统将自动启用 HTTPS。" -ForegroundColor Yellow
        exit
    }
    Write-Host "OpenSSL 执行失败，尝试其他方法..." -ForegroundColor Yellow
}

# Method 2: Use PowerShell's certificate generation
Write-Host "使用 PowerShell 生成证书..." -ForegroundColor Green

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "注意: 以管理员身份运行可获得更好的兼容性。" -ForegroundColor Yellow
}

try {
    # Generate using .NET
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        "CN=localhost",
        $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    # Add SAN extension for localhost and IP
    $san = [System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $false)
    $req.CertificateExtensions.Add($san)

    $subjectAlternativeName = [System.Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder]::new()
    $subjectAlternativeName.AddDnsName("localhost")
    $subjectAlternativeName.AddDnsName("127.0.0.1")
    try {
        $hostName = [System.Net.Dns]::GetHostEntry("").HostName
        $subjectAlternativeName.AddDnsName($hostName)
    } catch {}
    $req.CertificateExtensions.Add($subjectAlternativeName.Build())

    $cert = $req.CreateSelfSigned(
        [System.DateTimeOffset]::Now,
        [System.DateTimeOffset]::Now.AddDays(365)
    )

    # Export cert and key
    $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $certPem = "-----BEGIN CERTIFICATE-----`n"
    $certPem += [System.Convert]::ToBase64String($certBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
    $certPem += "`n-----END CERTIFICATE-----"

    $keyBytes = $rsa.ExportRSAPrivateKey()
    $keyPem = "-----BEGIN PRIVATE KEY-----`n"
    $keyPem += [System.Convert]::ToBase64String($keyBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
    $keyPem += "`n-----END PRIVATE KEY-----"

    Set-Content -Path $certFile -Value $certPem -Encoding ASCII
    Set-Content -Path $keyFile -Value $keyPem -Encoding ASCII

    Write-Host "证书生成成功！" -ForegroundColor Green
    Write-Host "  cert.pem (证书文件)" -ForegroundColor Gray
    Write-Host "  key.pem  (私钥文件)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "重启 start.bat 后，系统将自动启用 HTTPS。" -ForegroundColor Yellow
    Write-Host "局域网其他电脑可通过 https://你的IP地址:5000 访问" -ForegroundColor Yellow
    Write-Host "注意: 自签名证书会被浏览器提示不安全，选择"继续前往"即可。" -ForegroundColor Yellow
}
catch {
    Write-Host "证书生成失败: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "请尝试安装 OpenSSL (https://slproweb.com/products/Win32OpenSSL.html)" -ForegroundColor Yellow
    Write-Host "然后重新运行此脚本。" -ForegroundColor Yellow
    pause
}
