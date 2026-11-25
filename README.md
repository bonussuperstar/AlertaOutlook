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
* **Suport multimonitor opcional:** Permet mostrar l'avís simultàniament a tots els monitors connectats (desactivat per defecte). Tancar la finestra en qualsevol monitor tanca automàticament totes les altres.
* **Icona dinàmica a la safata (`NotifyIcon`):** Canvia d'estat visual entre repòs, alerta activa i mode pausat, amb control d'instància única (`Mutex`) per evitar processos duplicats.
* **Personalització visual completa:**
  * **Tema dels avisos:** **Automàtic (Sistema)** per defecte (detecta el mode clar/fosc de Windows al registre), o selecció manual **Fosc** / **Clar**.
  * **Color de destacat:** 5 colors predefinits o selector lliure (`ColorDialog`), aplicat a la franja superior, capçalera, botó **Acceptar** i icona d'alerta.
  * **Estil de la icona:** Sobre de correu vectorial, cercle (`M` / `!`) o parell de fitxers `.ico` personalitzats (repòs i alerta).
  * **Posició i so:** Centre de la pantalla o cantonada inferior dreta, amb so d'avís activable o desactivable.
* **Persistència a `%APPDATA%`:** Desa automàticament les preferències a `%APPDATA%\AlertaOutlook\config.json`, amb opcions al menú per obrir el fitxer o restablir els valors per defecte.
* **Consciència Per-Monitor V2 DPI:** Escalat nítid en pantalles d'alta densitat i configuracions multimonitor.

---

## Requisits

* **Windows** amb **PowerShell 5.1+** o **PowerShell 7+**.
* **Outlook clàssic d'escriptori** obert i amb sessió iniciada (l'objecte COM `Outlook.Application` requereix el client clàssic; el *Nou Outlook* basat en web no exposa l'API COM local).

---

## Instal·lació i ús

### Execució com a script (`.ps1`)
1. Desa `AlertaOutlookTray.ps1` en una carpeta local (per exemple, `C:\Scripts\AlertaOutlook\AlertaOutlookTray.ps1`).
2. Crea un accés directe (a l'escriptori o a `shell:startup` per iniciar-lo amb Windows) amb aquesta ordre:

```text
powershell.exe -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Scripts\AlertaOutlook\AlertaOutlookTray.ps1"

```

### Compilació a executable (`.exe`) amb `ps2exe`

Abans de compilar, assegura't que `AlertaOutlookTray.ps1` estigui desat en **UTF-8 amb BOM** perquè els caràcters accentuats es mostrin correctament en Windows PowerShell 5.1.

```powershell
Invoke-PS2EXE -InputFile ".\AlertaOutlookTray.ps1" `
              -OutputFile ".\AlertaOutlookTray.exe" `
              -noConsole `
              -STA `
              -noOutput `
              -noError `
              -title "AlertaOutlookTray" `
              -product "AlertaOutlookTray" `
              -version "1.0.0.0"

```

> **Important:**
> * **`-STA` és obligatori** perquè funcionin `ColorDialog`, `OpenFileDialog` i la connexió COM amb Outlook.
> * **No afegeixis `-requireAdmin`:** si l'executable corre elevat i Outlook corre com a usuari estàndard, Windows bloqueja la comunicació COM entre tots dos processos.
> 
> 

---

## Opcions del menú de la safata

Fent clic dret sobre la icona a la barra de tasques:

* **Provar avís ara:** Mostra una finestra de prova immediata per validar tema, color, posició i monitors.
* **Pausar avisos:** Silencia temporalment les alertes i canvia la icona a gris.
* **Tema dels avisos:** Tria entre **Automàtic (Sistema)**, **Fosc** o **Clar**.
* **Color de destacat:** Tria entre colors predefinits (*Vermell alerta*, *Taronja ambre*, *Blau Outlook*, *Verd maragda*, *Lila intens*) o *Tria color personalitzat...*.
* **Estil de la icona:** Selecciona *Sobre de correu*, *Cercle (M / !)* o *Carregar fitxers .ico...*.
* **Opcions d'avís:**
* *Reproduir so en arribar*
* *Mostrar a tots els monitors* (desactivat per defecte)
* *Posició: Centre pantalla* / *Cantonada inferior dreta*


* **Configuració:**
* **Obrir config.json:** Obre el fitxer de configuració al Bloc de notes.
* **Restablir valors per defecte:** Restaura la configuració inicial prèvia confirmació.


* **Sortir:** Atura el temporitzador, tanca totes les finestres d'avís obertes i neteja la icona de la safata.

---

## Fitxer de configuració

La configuració es crea i s'actualitza automàticament a:

```text
%APPDATA%\AlertaOutlook\config.json

```

Exemple de contingut per defecte:

```json
{
  "Theme": "System",
  "AccentHex": "#D93025",
  "IconStyle": "Envelope",
  "CustomIdleIcon": "",
  "CustomAlertIcon": "",
  "IntervalSec": 10,
  "PlaySound": true,
  "Position": "Center",
  "AllMonitors": false
}

```
