# AlertaOutlookTray

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B%20%7C%207%2B-5391FE?logo=powershell&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Outlook](https://img.shields.io/badge/Outlook-Classic%20Desktop%20(COM)-0F6CBD?logo=microsoftoutlook&logoColor=white)
![UI](https://img.shields.io/badge/UI-Windows%20Forms%20%7C%20Per--Monitor%20V2%20DPI-00A4EF)
![Config](https://img.shields.io/badge/Config-JSON%20(%25APPDATA%25)-2EA44F)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

Monitor de safata del sistema (**System Tray**) escrit en **PowerShell** amb **Windows Forms** que vigila la safata d'entrada de l'**Outlook clàssic d'escriptori** i mostra una finestra emergent persistent (`TopMost`) davant de totes les aplicacions quan arriba un correu nou sense llegir.

---

## Característiques principals

* **Finestra persistent davant de tot (`TopMost`):** Es manté visible fins que es tanca amb **Acceptar** o `Intro`, evitant perdre avisos quan no ets davant de la pantalla o treballes en altres finestres.
* **Actualització en viu mentre l'avís és obert:** Si arriben més correus mentre la finestra és activa, s'incrementa el comptador (`N CORREUS NOUS PENDENTS - ÚLTIM (...)`) i s'actualitzen remitent, assumpte i hora sense obrir finestres duplicades.
* **Icona dinàmica a la safata (`NotifyIcon`):** Canvia d'estat visual entre repòs, alerta activa i mode pausat, amb menú contextual per gestionar l'aplicació.
* **Personalització visual completa:**
  * **Tema dels avisos:** Fosc (`#1C1C1E`) o Clar (`#F8F9FA`).
  * **Color de destacat:** 5 colors predefinits o selector lliure (`ColorDialog`), aplicat a la franja superior, capçalera, botó **Acceptar** i icona d'alerta.
  * **Estil de la icona:** Sobre de correu vectorial, cercle (`M` / `!`) o parell de fitxers `.ico` personalitzats (repòs i alerta).
  * **Posició i so:** Centre de la pantalla o cantonada inferior dreta, amb so d'avís activable o desactivable.
* **Persistència a `%APPDATA%`:** Desa automàticament les preferències a `%APPDATA%\AlertaOutlook\config.json`, amb opcions al menú per obrir el fitxer o restablir els valors per defecte.
* **Consciència Per-Monitor V2 DPI:** Escalat nítid en pantalles d'alta densitat i configuracions multimonitor.

---

## Requisits

* **Windows** amb **PowerShell** i **Outlook clàssic d'escriptori** (l'objecte COM `Outlook.Application` requereix el client d'escriptori clàssic; el *Nou Outlook* basat en web no exposa l'API COM local).
* Outlook obert i amb sessió iniciada perquè la comprovació MAPI pugui accedir a la safata d'entrada (`olFolderInbox`).

---

## Instal·lació i ús

1. Desa l'script com a `AlertaOutlookTray.ps1` en una carpeta local (per exemple, `C:\Scripts\AlertaOutlook\AlertaOutlookTray.ps1`).
2. Per executar-lo **en segon pla i sense finestra de consola**, crea un accés directe (a l'escriptori o a `shell:startup` per arrencar amb sessió) amb aquesta ordre:

```text
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Scripts\AlertaOutlook\AlertaOutlookTray.ps1"

```

---

## Opcions del menú de la safata

Fent clic dret sobre la icona a la barra de tasques:

* **Provar avís ara:** Mostra una finestra de prova immediata per validar tema, color i posició.
* **Pausar avisos:** Silencia temporalment les alertes i canvia la icona a gris.
* **Tema dels avisos:** Alterna entre **Fosc** i **Clar**.
* **Color de destacat:** Tria entre colors predefinits (*Vermell alerta*, *Taronja ambre*, *Blau Outlook*, *Verd maragda*, *Lila intens*) o *Tria color personalitzat...*.
* **Estil de la icona:** Selecciona *Sobre de correu*, *Cercle (M / !)* o *Carregar fitxers .ico...*.
* **Opcions d'avís:** Activa/desactiva *Reproduir so en arribar* i canvia entre *Centre pantalla* i *Cantonada inferior dreta*.
* **Configuració:**
* **Obrir config.json:** Obre el fitxer de configuració a `%APPDATA%\AlertaOutlook\config.json`.
* **Restablir valors per defecte:** Restaura la configuració inicial prèvia confirmació.


* **Sortir:** Atura el temporitzador, tanca la finestra d'avís si és oberta i neteja la icona de la safata.

---

## Fitxer de configuració

La configuració es crea i s'actualitza automàticament a:

```text
%APPDATA%\AlertaOutlook\config.json

```

Exemple de contingut per defecte:

```json
{
  "Theme": "Dark",
  "AccentHex": "#D93025",
  "IconStyle": "Envelope",
  "CustomIdleIcon": "",
  "CustomAlertIcon": "",
  "IntervalSec": 10,
  "PlaySound": true,
  "Position": "Center"
}

```
