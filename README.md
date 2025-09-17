# ADQueryTool

Active Directory'den kullanıcı ve bilgisayar bilgilerini sorgulayan PowerShell script'i. Windows Forms GUI ile çalışır.

## Özellikler
- Dinamik sorgu şifre değiştirmemiş kullanıcılar.
- AdminCount=1/0 kullanıcılar.
- Dinamik sorgu login olmamış bilgisayarlar.
- OU filtresi ve credential desteği.
- Bilgisayar kapatma ve NetBIOS devre dışı bırakma.
- CSV/HTML çıktı ve HTML önizleme.

## Kurulum
1. PowerShell 5+ ve ActiveDirectory modülü gerekli.
2. Yönetici haklarıyla çalıştırın: `.\ADQueryTool.ps1`.

## Kullanım
- Domain adı girin (örneğin, quicksigorta.local).
- Gerekirse credential ekleyin.
- OU seçip bilgisayarları yükleyin, sorguları çalıştırın.

## Ön Koşullar
```powershell
Install-Module ActiveDirectory

**Uyarı**: Script, AD'de hassas işlemler (kapatma, NetBIOS) yapar. Test ortamında deneyin.
