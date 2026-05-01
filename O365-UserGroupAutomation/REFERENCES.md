# Referans GitHub Repolari

Bu klasordeki kodun gelistirilmesinde / karsilastirmasinda referans alinan
acik kaynak Office 365 / Microsoft 365 PowerShell projeleri. Hicbir kod
dogrudan kopyalanmadi; gelistirme srasinda kullanilmasi onerilen olgunlasmis
repolar listelenmistir. Kullanicinin GitHub hesabina **fork** edilip
gerektikce buradan adapte edilmesi onerilir.

> **Lisans uyarisi:** Bu repolarin her birinin kendi lisansi vardir. Buradan
> kod adapte ederken orijinal telif ve lisans metnini, kaynak commit
> hash'ini ve repo URL'ini koruyarak adapte edilen dosyaya basligini ekleyin.

## Aday repolar

### 1. 12Knocksinna/Office365itpros
- URL: https://github.com/12Knocksinna/Office365itpros
- Icerik: 380+ uretime yakin PowerShell script (Graph + Exchange Online).
- Bizim icin onemli dosyalar:
  - `Update-DynamicM365GroupAzureAutomation.PS1` — Azure Automation runbook
    olarak dinamik grup uyeligi guncelleme. `Sync-UserGroupMembership.ps1`
    icin esik karsilastirma.
  - `ReportMembershipM365Group.PS1` — Grup uyelik raporu.
    `Get-GroupReport.ps1` icin ek alan/feature kaynagi.
  - `Report-ExpiringPasswords.PS1` — Genisletilmis raporlama ornegi.

### 2. microsoft/Microsoft365DSC
- URL: https://github.com/microsoft/Microsoft365DSC
- Icerik: Resmi Microsoft projesi. Tenant konfigurasyonunu deklaratif olarak
  yonetir, drift detection saglar.
- Bizim icin onemli noktalar:
  - `MSFT_AADGroup` ve `MSFT_AADUser` DSC kaynaklari — kural bazli yonetimi
    DSC ile yapmak istersek baz alinabilir.
  - CI/CD pipeline'a tasimak icin orneklenen yapilar.

### 3. jhoneill/MsftGraph
- URL: https://github.com/jhoneill/MsftGraph
- Icerik: Microsoft Graph API uzerine yazilmis ust seviye PowerShell modulu.
- Bizim icin onemli noktalar:
  - `Get-GraphUser` / `Get-GraphGroup` gibi sade wrapper'lar; mevcut
    Get-MgUser cagrilarimizi sadelestirmek icin opsiyonel bagimlilik.

### 4. spiddeer/M365-Scripts
- URL: https://github.com/spiddeer/M365-Scripts
- Icerik: SPO, Teams, kullanici/grup yonetimi icin pratik scriptler.
- Bizim icin onemli noktalar: Kucuk olcekli sysadmin senaryolari icin
  hazir reference.

### 5. CurtisSlone/M365DSC
- URL: https://github.com/CurtisSlone/M365DSC
- Icerik: DevSecOps + GitHub Actions ile M365DSC drift detection ornek
  uygulamasi.
- Bizim icin onemli noktalar:
  - `.github/workflows/*.yml` sablonlari — bu repoda CI/CD baslatmak icin
    referans.

## Onerilen fork komutlari (kullanici tarafi)

GitHub MCP entegrasyonu yalnizca `deepdarbe/office365` reposuna kisitli
oldugundan asagidaki forklar manuel veya yerel `gh` ile yapilmalidir:

```bash
gh repo fork 12Knocksinna/Office365itpros --clone=false
gh repo fork microsoft/Microsoft365DSC    --clone=false
gh repo fork jhoneill/MsftGraph           --clone=false
gh repo fork spiddeer/M365-Scripts        --clone=false
gh repo fork CurtisSlone/M365DSC          --clone=false
```

Fork tamamlandiktan sonra bu dosyadaki repo URL'lerini kendi fork
URL'leriniz ile guncelleyin (orn. `https://github.com/<sizin-org>/Office365itpros`).

## Adaptasyon kurallari

1. **Asla** orijinal repodaki tenant verilerini, mail adreslerini, GUID'leri
   bizim repoya tasimayin (zaten orijinal repolarda da bunlar yok ama yine
   de check edin).
2. Adapte edilen her dosyanin basina kaynak ve lisans yorum bloku ekleyin:
   ```
   # Source: https://github.com/12Knocksinna/Office365itpros/blob/<sha>/<file>
   # License: MIT (c) 2024 Tony Redmond / Office 365 for IT Pros
   # Adapted: <tarih> — <kisa not>
   ```
3. Bu repodaki `.gitignore` (`data/`, `*.csv`, `*.log`, `group-mapping.json`,
   `known-users.json`) adapte edilen scriptlerin ciktilarini da kapsadigi
   icin tenant verisi sizmaz.

## Backlog (paralel arastirma sonrasi)

- [ ] `Update-DynamicM365GroupAzureAutomation.PS1`'i incele, bizim
      `Sync-UserGroupMembership.ps1` ile feature parity karsilastirmasi yap.
- [ ] Microsoft365DSC ile drift detection POC.
- [ ] CurtisSlone/M365DSC workflow ornegine bakarak bu repoya `.github/workflows/`
      altinda lint + scriptanalyzer pipeline ekle.
- [ ] jhoneill/MsftGraph wrapper'larinin bizim koda deger katip katmadigini
      degerlendir (genelde gerekmez; saf Microsoft.Graph yeterli).
