# Bokføringsoppdrag – app for Teams og SharePoint (SPFx)

Felles liste over bokføringsoppdrag (klient, oppgavetype og periode) som vises som egen app i Teams
og som web-del på SharePoint-sider. Alt som legges inn lagres i en skjult SharePoint-liste
(`Oppdrag`), så alle i firmaet ser og redigerer den samme listen.

Ferdig pakke: **`oppdrag.sppkg`** (i denne mappen).

## Installere (Microsoft 365-administrator, én gang)

1. **Velg område for dataene.** Lag et vanlig SharePoint-område (for eksempel *Regnskap*) eller bruk et
   eksisterende. Alle som skal bruke appen må ha **Rediger**-tilgang (medlem) på området.
2. **Last opp pakken.** SharePoint admin center → *Mer funksjoner* → *Apper* → *Åpne* → *Last opp* →
   velg `oppdrag.sppkg`. Kryss av for **«Gjør denne appen tilgjengelig for alle områder i organisasjonen»**
   og trykk *Aktiver*.
3. **Legg den i Teams.** Merk appen i app-katalogen og trykk **«Legg til i Teams»** (Sync to Teams).
   Den dukker opp som *Bokføringsoppdrag* under *Bygd for organisasjonen din* i Teams.
4. **Pek appen mot dataområdet** (anbefalt, gjelder alle brukere). I SharePoint Online Management Shell / PnP:

   ```powershell
   # PnP PowerShell, koblet til app-katalogområdet
   Set-PnPStorageEntity -Key OppdragDataSite -Value "https://<firma>.sharepoint.com/sites/regnskap" -Description "Område for Bokføringsoppdrag"
   ```

   Uten denne innstillingen bruker appen området den kjører på: som Teams-fane i et team er det teamets
   SharePoint-område, som personlig app er det rotområdet.
5. **Åpne appen første gang som eier av dataområdet.** Da opprettes den skjulte listen `Oppdrag`
   automatisk. Vanlige medlemmer kan ikke opprette lister og får en tydelig melding om dette.

## Bruke

* **Personlig app:** I Teams → *Apper* → *Bygd for organisasjonen din* → *Bokføringsoppdrag* → *Legg til*.
  Fest den i venstremenyen.
* **Fane i et team/kanal:** *+* øverst i kanalen → *Bokføringsoppdrag*.
* **SharePoint-side:** Rediger side → *+* → *Bokføringsoppdrag*. Under web-delens innstillinger kan du
  eventuelt overstyre hvilket område listen hentes fra.

Listen oppdateres automatisk hvert minutt og når du går tilbake til fanen. Hvis to personer endrer samme
oppdrag samtidig, får den siste en melding i stedet for å overskrive den andres endring.
Mørk modus følger Teams-temaet.

## Utvikle og bygge

Krever Node 22.

```bash
npm install
npm run build          # bygger og lager sharepoint/solution/oppdrag.sppkg
cp sharepoint/solution/oppdrag.sppkg .
```

`src/webparts/oppdrag/core/` inneholder brukergrensesnittet (`ui.ts`) og lagringen mot SharePoint
(`store.ts`) uten avhengigheter til SPFx. De kan prøves lokalt mot en simulert SharePoint:

```bash
npm run build && npx http-server . -p 8765    # åpne http://localhost:8765/dev/  (?mork for mørk modus)
```

Ved ny versjon: øk `version` i `config/package-solution.json`, bygg, og last opp pakken på nytt i
app-katalogen (erstatt eksisterende).
