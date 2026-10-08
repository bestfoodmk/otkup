# Откуп на терен — самостојна апликација (PWA) без Claude

Теренската апликација за внес на откупено млеко, како веб-апликација што се инсталира на iPad од почетниот екран, се отвора и без интернет, а податоците ги праќа во две SharePoint листи во Microsoft 365 на Ѓоргиеви.

## Што има во папката

| Датотека | Намена |
|---|---|
| `index.html`, `app.js` | Теренската апликација за откупувачите |
| `zaednicko.js` | Најава со Microsoft (MSAL) и пристап до SharePoint преку Microsoft Graph |
| `config.js` | **Единствената датотека што се пополнува при инсталација** |
| `sw.js`, `manifest.webmanifest`, икони | Инсталација на почетен екран и работа без интернет |
| `setup.html` | Еднократно: ги создава листите во SharePoint и ги увезува кооперантите од JSON |
| `kancelarija.html` | Канцеларија: преглед по ден, потврда на внесови, CSV, следна слободна шифра |
| `lib/` | jsQR (читање QR од слика) и msal-browser, локални копии за офлајн |

## Инсталација (еднократно, околу 1 час)

### 1. Регистрација на апликацијата во Entra ID (Azure AD)
Во [entra.microsoft.com](https://entra.microsoft.com) → Applications → App registrations → New registration:
- Name: `Otkup na teren`
- Supported account types: *Accounts in this organizational directory only*
- Redirect URI: платформа **Single-page application (SPA)**, адреса = онаму каде ќе биде хостирана апликацијата, `https://bestfoodmk.github.io/otkup/`. Додади ги и `…/setup.html` и `…/kancelarija.html`.

Потоа:
- Overview → препиши **Application (client) ID** и **Directory (tenant) ID** во `config.js`.
- API permissions → Add → Microsoft Graph → Delegated → `Sites.ReadWrite.All` и `User.Read` → **Grant admin consent**.
- Authentication → Allow public client flows: не е потребно. Implicit grant: оставете исклучено (MSAL користи PKCE).

### 2. Хостирање
Било кој статички хостинг со HTTPS. Најлесно:
- **Azure Static Web Apps (Free)**: Create → Source „Other“ → по креирањето, во Overview → „Manage deployment token“, и прикачете ја папката со SWA CLI, или поврзете GitHub репозиториум со оваа папка. Или
- **GitHub Pages**: репозиториум со овие датотеки → Settings → Pages.

Адресата на хостингот мора да биде иста како Redirect URI од чекор 1. Пополни `config.js` (tenantId, clientId, site).

### 3. SharePoint листи
Отвори `https://<адреса>/setup.html`, најави се со сметка што смее да создава листи на `bfgj.sharepoint.com/sites/mlekarnica`, притисни **Создај ги листите**, па **Увези** со `../Otkup_sifri_2026-10-05/spisok_site_2026-10-05.json`. Листите:
- `Kooperanti`: Title (име), Sifra (единствена), Vid (kravjo/ovco/kozjo), Reon, ReonBr, Mesto, Otk, Aktiven, Prethodna
- `OtkupVnesovi`: Title (име), Sifra, Vid, Litri, Temp, Zabeleska, Den, Reon, ReonIme, Koj, Vreme, LokalenId (единствен, спречува дупли), Status (Примено / Потврдено / Одбиено)

Нов кооперант или нова шифра се внесува директно во листата `Kooperanti` во SharePoint; таблетите ја добиваат со „Освежи го списокот“.

### 4. Сметка за таблетите
Една сметка во Microsoft 365 за сите таблети: **milk@ti.com.mk**, со дозвола **Edit** само на сајтот `mlekarnica` (или само на двете листи). Не користете администраторска сметка.

### 5. На секој iPad
1. Safari → отвори ја адресата → Share → **Add to Home Screen**.
2. Отвори ја од иконата → јазиче **Прати** → **Најави се** со сметката за таблети.
3. **Освежи го списокот**, па впиши име на откупувачот и реон → **Почни со внес**.

## Работа без сигнал
- Апликацијата се отвора од иконата и без интернет (сè е кеширано на уредот).
- Внесовите се чуваат на таблетот и се праќаат сами кога ќе има мрежа, или со **Прати**.
- Списокот на кооперанти се чува локално од последното освежување.
- Најавата важи долго; ако истече, бара повторна најава само кога има интернет.

## Ажурирање на апликацијата
Промени ги датотеките, зголеми `verzija` во `config.js` и `VERZIJA` во `sw.js`, прикачи повторно. Таблетите ја повлекуваат новата верзија при следното отворање со интернет.

## Тестирање локално
```
cd Otkup_na_teren_PWA && python3 -m http.server 8080
```
Најавата работи само ако `http://localhost:8080/` е додаден како SPA Redirect URI во регистрацијата.
