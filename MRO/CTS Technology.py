import os
import pdfplumber
import pandas as pd
import re


def remove_polish_chars(text):
    if text is None:
        return text
    replacements = {
        "ą": "a",
        "ć": "c",
        "ę": "e",
        "ł": "l",
        "ń": "n",
        "ó": "o",
        "ś": "s",
        "ź": "z",
        "ż": "z",
        "Ą": "A",
        "Ć": "C",
        "Ę": "E",
        "Ł": "L",
        "Ń": "N",
        "Ó": "O",
        "Ś": "S",
        "Ź": "Z",
        "Ż": "Z",
    }
    value = str(text)
    for old_char, new_char in replacements.items():
        value = value.replace(old_char, new_char)
    return value

def process_all_pdfs_in_folder(katalog_z_pdfami, plik_wyjsciowy):
    wszystkie_dane = []

    if not os.path.exists(katalog_z_pdfami):
        print(f"❌ Podany folder nie istnieje: {katalog_z_pdfami}")
        return

    licznik_plikow = 0
    for nazwa_pliku in os.listdir(katalog_z_pdfami):
        if not nazwa_pliku.lower().endswith('.pdf'):
            continue
            
        licznik_plikow += 1
        pelen_odnosnik = os.path.join(katalog_z_pdfami, nazwa_pliku)
        znaleziono_w_pliku = 0
        print(f"📄 Przetwarzam plik CTS: {nazwa_pliku}")
        
        try:
            with pdfplumber.open(pelen_odnosnik) as pdf:
                # Czytamy tylko pierwszą stronę na potrzeby testu (zmień na pdf.pages jeśli oferty mają wiele stron)
                tekst = pdf.pages[0].extract_text()
                
                if tekst:
                    # Pobieramy termin realizacji na całą ofertę
                    termin_realizacji = "Brak informacji"
                    termin_match = re.search(r'Termin realizacji:\s*([^\n\r]+)', tekst, re.IGNORECASE)
                    if termin_match:
                        termin_realizacji = termin_match.group(1).strip()
                    termin_realizacji = remove_polish_chars(termin_realizacji)
                        
                    # Rozbijamy cały tekst na linijki
                    for linia in tekst.split('\n'):
                        # KLUCZOWE: Usuwamy wszystkie spacje na początku i końcu linii
                        linia = linia.strip() 
                        
                        # Szukamy linii, która zaczyna się od jakiejś cyfry (np. 1, 2 - jako Lp.) 
                        # i zawiera kwotę z przecinkiem i dwiema cyframi po przecinku
                        kwota_match = re.search(r'(\d+[\s\d]*,\d{2})\s*([A-Za-z\$€]+)?', linia)
                        if re.search(r'^\d', linia) and kwota_match:
                            
                            # 2. Wyciągamy symbol (wzór: ciąg znaków, kropka, ciąg znaków i myślników - np. 01168.M04-14-04)
                            symbol_match = re.search(r'([A-Za-z0-9]+\.[A-Za-z0-9\-]+)', linia)
                            
                            if symbol_match:
                                symbol = remove_polish_chars(symbol_match.group(1).strip())
                                cena = kwota_match.group(1).strip() # Sama kwota
                                waluta = kwota_match.group(2) if kwota_match.group(2) else ""
                                naglowek_ceny = remove_polish_chars(
                                    f"Cena jednostkowa netto {waluta}".strip()
                                )
                                
                                # Nazwa to to, co znajduje się przed symbolem
                                poczatek_linii = linia[:symbol_match.start()].strip()
                                # Usuwamy numer "Lp." z samego początku (nawet jeśli brakuje spacji)
                                nazwa = re.sub(r'^\d+\s*', '', poczatek_linii).strip()
                                nazwa = remove_polish_chars(nazwa)
                                
                                wszystkie_dane.append({
                                    "Plik zrodlowy": remove_polish_chars(nazwa_pliku),
                                    "Nazwa towaru/ uslugi": nazwa,
                                    "Symbol": symbol,
                                    naglowek_ceny: cena,
                                    "Termin realizacji": termin_realizacji,
                                })
                                znaleziono_w_pliku += 1
                                
        except Exception as e:
            print(f"Błąd krytyczny przy pliku {nazwa_pliku}: {e}")

        # -------- DIAGNOSTYKA --------
        if znaleziono_w_pliku == 0:
            print(f" ❌ Nic nie znaleziono w pliku: {nazwa_pliku}")
        else:
            print(f" ✅ Wyciągnięto {znaleziono_w_pliku} pozycji z: {nazwa_pliku}")

    if licznik_plikow == 0:
        print("⚠️ Nie znaleziono plików PDF w folderze CTS!")
        return

    # Zapis do pliku
    if wszystkie_dane:
        df = pd.DataFrame(wszystkie_dane)
        df.to_excel(plik_wyjsciowy, index=False)
        print(f"\n🎉 SUKCES! Wygenerowano plik: {plik_wyjsciowy} z {len(df)} pozycjami dla CTS.")
    else:
        print("\n⚠️ Skrypt CTS zakończył pracę, ale lista jest pusta.")

if __name__ == '__main__':
    # Kod testowy, jeśli skrypt odpalony bezpośrednio
    process_all_pdfs_in_folder("C:/Users/dawipude/Downloads/Oferty", "zbiorcze_dane_z_ofert.xlsx")