# Generator testów

Proste narzędzie z graficznym interfejsem (GUI) do generowania wielu zrandomizowanych wersji testu z jednego szablonu `.docx` (pytania w tabeli w Wordzie) wraz z plikami kluczy odpowiedzi.

## Skąd pobrać
Gotowe pliki wykonywalne (EXE dla Windows oraz binarki dla Linuksa) są publikowane w sekcji [Releases](https://github.com/jmgraa/docx-test-generator/releases) tego repozytorium.  
Wystarczy pobrać odpowiedni plik dla swojego systemu operacyjnego i uruchomić go – nie ma potrzeby instalowania Pythona ani dodatkowych bibliotek.

## Jak korzystać z programu

1. Uruchom pobrany plik – otworzy się okno aplikacji.
2. W otwartym oknie:
   - **Liczba kopii**: podaj, ile różnych wersji testu chcesz wygenerować.
   - **Nagłówek testu**: wpisz tytuł/nagłówek testu.
   - **Data testu**: potwierdź lub zmień datę testu.
   - **Wstęp do testu**: w razie potrzeby zmodyfikuj tekst wstępu.
   - Kliknij **Wybierz plik Word** i wskaż źródłowy plik `.docx` z pytaniami (w formie tabeli).
   - Kliknij **Wybierz folder wyjściowy** i wybierz katalog, w którym mają zostać zapisane wygenerowane testy oraz klucze.
   - Kliknij **Generuj pliki**, aby rozpocząć generowanie. Postęp będzie wyświetlany na dole okna.

Wygenerowane pliki zostaną zapisane w wybranym katalogu wyjściowym w podkatalogach przeznaczonych na pliki testów i pliki z kluczami odpowiedzi.

## Ważna uwaga dotycząca obrazów w pliku DOCX

Jeżeli Twoje pytania zawierają **obrazy umieszczone w tabeli w pliku DOCX**, wszystkie te obrazy **muszą mieć zablokowany (jawnie ustawiony) rozmiar – szerokość i wysokość – w edytorze Word**.

Zapewnia to, że:
- **we wszystkich wygenerowanych wersjach testu obrazy będą miały ten sam rozmiar**, oraz  
- Word nie będzie automatycznie zmieniał ich wielkości między plikiem źródłowym a wynikowymi plikami `.docx`.

## Korzystanie z Pythona (dla użytkowników zaawansowanych)

### Przygotowanie wirtualnego środowiska Pythona

Program wykorzystuje bibliotekę `tkinter`.  
Na Windowsie `tkinter` jest zwykle dostępny razem z Pythonem, natomiast **w wielu dystrybucjach Linuksa trzeba go najpierw doinstalować jako pakiet systemowy** (np. `python-tkinter`, `python3-tk`, itp. – nazwa zależy od dystrybucji).  
Po zapewnieniu obecności `tkinter` w systemie, pozostałe zależności można zainstalować we wirtualnym środowisku z pliku `requirements.txt`.

W katalogu głównym repozytorium:

```bash
# 1. Utwórz wirtualne środowisko
python -m venv .venv

# 2. Aktywuj je
.venv\Scripts\Activate.ps1 # Windows

source .venv/bin/activate  # Linux / macOS

# 3. Zainstaluj zależności
pip install -r requirements.txt

# 4. Uruchom aplikację z GUI z katalogu projektu:
python gui.py
```

### Budowanie samodzielnego pliku wykonywalnego (PyInstaller)

Do zbudowania pliku wykonywalnego potrzebny jest **PyInstaller**.  
Jeśli nie jest zainstalowany, dodaj go w aktywnym wirtualnym środowisku:

```bash
pip install pyinstaller
```

Następnie, aby zbudować pojedynczy plik wykonywalny z GUI, uruchom:

```bash
pyinstaller gui.py --name docx-test-generator --collect-all docx --onefile --noconsole
```

Po zakończeniu procesu plik wykonywalny będzie dostępny w katalogu `dist/` pod nazwą `docx-test-generator`.

