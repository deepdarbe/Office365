# O365 User & Group Automation

Office 365 / Microsoft 365 tenant'inda **aktif kullanicilari cekme**, **yeni
kullanicilari kural bazli olarak gruplara otomatik ekleme** ve gruplari
**analiz / liste / duzenleme / raporlama** islemlerini yapan PowerShell
modulu. Microsoft Graph + Exchange Online PowerShell uzerinden calisir.

> **Onemli:** Gercek tenant verileri (kullanici listeleri, mapping, loglar)
> bu repoya **commitlenmez**. Hassas dosyalar `.gitignore` ile haric
> tutulmustur. Mapping dosyanizi `group-mapping.template.json`'dan
> kopyalayarak yerel olarak `group-mapping.json` adiyla olusturun.

## Gereksinimler

```powershell
Install-Module Microsoft.Graph         -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

PowerShell 7+ onerilir. Hesabinizin Graph API icin
`User.Read.All`, `Group.ReadWrite.All`, `GroupMember.ReadWrite.All`,
`Directory.Read.All` izinlerine onay vermesi gerekir.

## Hizli baslangic

```powershell
# 1) Modulleri yukle ve baglan (MFA destekli interaktif giris)
. .\Connect-O365Services.ps1
Connect-O365 -TenantId 'contoso.onmicrosoft.com'

# 2) Aktif (enabled + lisansli) kullanicilari listele ve CSV'ye yaz
.\Get-ActiveUsers.ps1 -ExportCsv

# 3) Mapping dosyasini olustur (sadece ilk kez)
Copy-Item .\group-mapping.template.json .\group-mapping.json
# group-mapping.json'i kendi kurallariniza gore duzenleyin

# 4) Yeni kullanicilari otomatik gruplara ekle (kuru calisma)
.\Sync-UserGroupMembership.ps1 -OnlyNewUsers -WhatIf

# 5) Gercek calistirma
.\Sync-UserGroupMembership.ps1 -OnlyNewUsers

# 6) Grup raporu (CSV + HTML)
.\Get-GroupReport.ps1 -Detailed -Html

# 7) Grup yonetimi
.\Manage-Groups.ps1 -Action Create -GroupName 'NBR BURSA' `
    -MailNickname 'nbr-bursa' -GroupType MailEnabledSecurity `
    -Description 'NBR Bursa toplu mail grubu'

.\Manage-Groups.ps1 -Action AddMember    -GroupName 'NBR BURSA' -UserUpn 'ali@contoso.com'
.\Manage-Groups.ps1 -Action RemoveMember -GroupName 'NBR BURSA' -UserUpn 'ali@contoso.com'
.\Manage-Groups.ps1 -Action BulkFromCsv  -CsvPath .\data\bulk.csv
```

## Mapping kurallari

`group-mapping.json` (template'den kopyalanir):

```json
{
  "defaultGroups": ["All Staff"],
  "rules": [
    { "name": "Bursa ofisi",
      "match": { "field": "Department", "operator": "equals", "value": "NBR Bursa" },
      "groups": [ "NBR BURSA" ] }
  ]
}
```

- **field**: `Department`, `OfficeLocation`, `JobTitle`, `City`, `Country`,
  `UserPrincipalName`, `Mail`, `DisplayName` vb. (Get-ActiveUsers ciktisindaki alanlar)
- **operator**: `equals`, `notEquals`, `startsWith`, `endsWith`, `contains`, `regex`
- **groups**: Eslesen kullanicinin uye olacagi grup display name'leri

`defaultGroups` her aktif kullaniciya uygulanir; `rules` ise eslestiginde ek
gruplara uyelik ekler.

## Yeni kullanici tespiti

`Sync-UserGroupMembership.ps1` `data/known-users.json` icinde her UPN'in ilk
gorulus tarihini saklar. `-OnlyNewUsers` parametresi ile sadece daha once
gorulmemis kullanicilara mapping uygular. Bu dosya commit edilmez.

Cron / Zamanlanmis Gorev / Azure Automation icinde dakikalik veya saatlik
calistirilabilir; her yeni kullanici otomatik olarak ilgili gruplara dahil
olur.

## Dosyalar

| Dosya | Aciklama |
|-------|----------|
| `Connect-O365Services.ps1`    | Graph + EXO baglanti yardimcisi |
| `Get-ActiveUsers.ps1`         | Aktif kullanicilari ceker, CSV uretir |
| `Sync-UserGroupMembership.ps1`| Mapping kurallari ile uyelik atar |
| `Get-GroupReport.ps1`         | Grup ozet / detay raporu (CSV/HTML) |
| `Manage-Groups.ps1`           | Grup olusturma/guncelleme/silme/uyelik |
| `group-mapping.template.json` | Mapping kural sablonu |
| `data/`                       | Cikti, log ve state (gitignore'lu) |

## Guvenlik notlari

- Mapping dosyasi, CSV ciktilari, log ve state dosyalari `data/` icine yazilir
  ve `.gitignore` ile haric tutulur. Hicbir musteri verisi GitHub'a gitmez.
- Servis hesabi yerine modern auth ile MFA li hesap kullanin; otomasyon icin
  Azure AD app registration + sertifika tabanli yetki onerilir.
- Toplu silme veya `BulkFromCsv` islemlerinden once `-WhatIf` ile dogrulayin.
