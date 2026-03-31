# ⏱ Clockify → Google Sheets

Skrypt Google Apps Script, który automatycznie pobiera wpisy czasu z Clockify i generuje raport godzinowy w Google Sheets — zgodny z firmowym szablonem rozliczeniowym.

---

## 📋 Spis treści

- [Wymagania](#-wymagania)
- [Instalacja krok po kroku](#-instalacja-krok-po-kroku)
- [Konfiguracja](#-konfiguracja)
- [Użytkowanie](#-użytkowanie)
- [Generowanie raportów na kolejne miesiące](#-generowanie-raportów-na-kolejne-miesiące)
- [Automatyczne odświeżanie (opcjonalnie)](#-automatyczne-odświeżanie-opcjonalnie)
- [Struktura raportu](#-struktura-raportu)
- [Rozwiązywanie problemów](#-rozwiązywanie-problemów)

---

## 🔧 Wymagania

- Konto [Clockify](https://clockify.me/) z zarejestrowanymi wpisami czasu
- Konto Google z dostępem do Google Sheets
- Klucz API Clockify
- Workspace ID oraz User ID z Clockify

---

## 🚀 Instalacja krok po kroku

### 1. Pobierz klucz API Clockify

1. Zaloguj się do [Clockify](https://app.clockify.me/)
2. Przejdź do **Manage API view** → [https://app.clockify.me/user/preferences#advanced](https://app.clockify.me/user/preferences#advanced)
3. Wygeneruj klucz i skopiuj go

### 2. Znajdź swoje Workspace ID i User ID

1. W Clockify podejrzysz sobie DevTools w zakladce network, przeklikaj UI clockify
2. W ktoryms requescie bedzie workspaceId i userId

Jesli nie udalo Ci sie uzyskac workspace i user ID patrz pkt. 5

### 3. Utwórz nowy arkusz Google Sheets

1. Przejdź do [Google Sheets](https://sheets.google.com/) i utwórz nowy arkusz
2. Nadaj mu nazwę, np. `Raport godzinowy - {miesiac} {rok}`

### 4. Otwórz edytor Apps Script

1. W arkuszu przejdź do **Rozszerzenia → Apps Script**
2. Usuń domyślny kod w pliku `Code.gs`
3. Wklej cały kod ze skryptu (`clockify-sheets.gs`)
4. Zapisz projekt (`Ctrl + S`)

### 5. Uzyskaj swoje User ID i Workspace ID ( jesli jeszcze nie masz )

1. Tymczasowo dodaj poniższą funkcję na końcu skryptu:

```javascript
function getMyUserId() {
  const response = callClockify(`${BASE_URL}/user`);
  const user = JSON.parse(response.getContentText());
  SpreadsheetApp.getUi().alert(`Your User ID: ${user.id}\nName: ${user.name}`);
}
```

i w funkcji `onOpen`

```javascript
    .addItem("Get User ID", "getMyUserId")
```

2. W edytorze Apps Script z listy rozwijanej wybierz funkcję `getMyUserId` i kliknij **Uruchom** (▶)
3. Google poprosi o autoryzację — zaakceptuj uprawnienia
4. Pojawi się alert z Twoim **User ID** — skopiuj go
5. Po skopiowaniu możesz usunąć tę funkcję ze skryptu
6. Po workspace ID skontaktuj sie z kims kto uzywa tego skryptu

---

## ⚙ Konfiguracja

W sekcji `CONFIG` na górze skryptu uzupełnij swoje dane:

```javascript
const CONFIG = {
  API_KEY: "twój-klucz-api", // Klucz API z Clockify
  WORKSPACE_ID: "twoje-workspace-id", // ID workspace'a
  USER_ID: "twoje-user-id", // ID użytkownika (z kroku 6)
  SHEET_NAME: "Time Logs", // Nazwa zakładki w arkuszu
  PAGE_SIZE: 200, // Maks. wpisów na stronę API

  // Dane do raportu
  NAME: "Jan Kowalski", // Twoje imię i nazwisko
  POSITION: "Developer", // Twoje stanowisko
  PROJECT: "", // Zostaw puste = projekt z Clockify
  // Wpisz nazwę = nadpisuje projekt z Clockify eq. "MArAdTech"
};
```

---

## 📊 Użytkowanie

1. Odśwież arkusz Google Sheets (lub zamknij i otwórz ponownie)
2. Na pasku menu pojawi się nowe menu **⏱ Clockify**
3. Dostępne opcje:
   - **Fetch Current Month** — pobiera wpisy z bieżącego miesiąca
   - **Fetch by Date Range** — pobiera wpisy z wybranego zakresu dat (format: `YYYY-MM-DD`)

---

## 📅 Generowanie raportów na kolejne miesiące

Skrypt jest przypisany do konkretnego arkusza. Aby wygenerować raport na nowy miesiąc **nie musisz ponownie konfigurować skryptu** — wystarczy skopiować arkusz:

1. Otwórz swój skonfigurowany arkusz (szablon)
2. Przejdź do **Plik → Utwórz kopię**
3. Nadaj nową nazwę, np. `Raport godzinowy - {month} {year}`
4. Otwórz skopiowany arkusz
5. Kliknij **⏱ Clockify → Fetch Current Month**

Kopia zawiera pełny skrypt z Twoją konfiguracją — wystarczy jedno kliknięcie, żeby wygenerować nowy raport.

> 💡 **Wskazówka:** Zachowaj jeden arkusz jako „szablon główny" i nigdy go nie usuwaj. Każdego miesiąca twórz z niego kopię.

---

## ❓ Rozwiązywanie problemów

| Problem                              | Rozwiązanie                                                                                |
| ------------------------------------ | ------------------------------------------------------------------------------------------ |
| Nie widzę menu **⏱ Clockify**        | Odśwież stronę arkusza lub zamknij i otwórz ponownie                                       |
| Błąd autoryzacji                     | Przy pierwszym uruchomieniu Google wymaga zgody — kliknij „Zaawansowane" → „Przejdź do..." |
| Brak wpisów z ostatnich dni miesiąca | Upewnij się, że używasz najnowszej wersji skryptu (daty są budowane w UTC)                 |
| `Clockify API error (401)`           | Sprawdź poprawność klucza API w CONFIG                                                     |
| `Clockify API error (403)`           | Sprawdź Workspace ID i User ID                                                             |
| Nowy arkusz nie ma menu              | Skrypt jest przypisany do arkusza — użyj **Plik → Utwórz kopię**                           |
