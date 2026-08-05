# EXE en Setup.exe bouwen – Fashion Image Renamer v5

## 1. Repository voorbereiden

1. Pak deze ZIP lokaal uit.
2. Maak een lege privé-repository op GitHub.
3. Upload de **inhoud** van de uitgepakte map naar de hoofdmap van de repository.
4. Controleer dat deze bestanden exact op deze plaats staan:

```text
.github/workflows/build-exe.yml
.github/workflows/build-setup.yml
app.py
requirements.txt
installer.iss
fotocategorieen.json
matchconfig.json
app_icon.ico
```

De `.yml`-bestanden mogen niet los in de hoofdmap staan.

## 2. Losse EXE bouwen

1. Open in GitHub het tabblad **Actions**.
2. Kies links **Build losse EXE**.
3. Klik **Run workflow** en nogmaals **Run workflow**.
4. Open de afgeronde run.
5. Download onder **Artifacts**: `FashionImageRenamer-v5-losse-EXE`.

Daarin staat:

```text
FashionImageRenamer_v5.exe
```

## 3. Setup.exe bouwen

1. Open **Actions**.
2. Kies links **Build Setup.exe**.
3. Klik **Run workflow**.
4. Download na afloop het artifact `FashionImageRenamer-v5-Setup`.

Daarin staat:

```text
FashionImageRenamer_v5_Setup.exe
```

## 4. Configuratie

`matchconfig.json` bepaalt:

- het scheidingsteken in de bestandsnaam;
- de gebruikte Excel-kolommen;
- exacte of bevat-matching.

De standaard is:

```text
vóór _  -> Artikelnr. Leverancier (exact)
na _    -> Externe omschrijving (bevat)
```

`fotocategorieen.json` bepaalt de nummering per soort beeld.
