# Importy modułów standardowych
import re
import os
import sys
import time
import ctypes
import msvcrt
import configparser
# Importy modułów zewnętrznych
import pandas as pd
from scipy.signal import savgol_filter

# Aby uruchomić program należy najpierw zainstalować pythona 3.11 lub nowszego i doinstalować trzy biblioteki:
# pip install pandas xlsxwriter scipy
# Program wczytuje wszystkie pliki z tribometru i generuje wykresy w pliku .xlsx
# Nazwa pliku + .txt lub .csv to nazwa wykresu
# Jeśli w nazwie pliku jest nawias to tekst w nim zawarty będzie opisem osi, dla title_from_text = 1
# Jeśli jest plik config.ini lub _config.ini to dane są wczytywane z niego najpierw
# Dla Rtec w nazwie musi być zawarta prędkość liniowa (np. 0.1m-s), a dla T11 dodatkowo obciążenie (np. 10N)

# Load config from file
def load_config(file_name):
    config = configparser.ConfigParser()
    config.read(file_name)
    settings = config["Settings"]
    return {
        "offset_raw": int(settings.get("offset_raw", 1)),
        "title_from_text": int(settings.get("title_from_text", 1)),
        "erase_peak": int(settings.get("erase_peak", 0)),
        "invert_peak": int(settings.get("invert_peak", 0)),
        "default_window_length_u": int(settings.get("default_window_length_u", 7)),
        "default_window_length_pd": int(settings.get("default_window_length_pd", 5)),
        "min_sample": int(settings.get("min_sample", 100)),
        "max_sample": int(settings.get("max_sample", 110)),
        "chart_lang": settings.get("chart_lang", "en")
    }

# Get variables from user
def ask_user_for_variables():
    pd_set = 3
    offset_raw, erase_peak, invert_peak = 1, 0, 0
    title_from_text = 1
    default_window_length_u, default_window_length_pd = 7, 5
    min_sample, max_sample = 100, 110
    chart_lang = "en"
    while True:
        offset_raw_input = input(f"Czy offsetować dane RAW zużycia liniowego? [1 - tak, 0 - nie] ({offset_raw}): ").strip() or '1'
        title_from_text_input = input(f"Czy nazwać serię danych tekstem z nawiasu nazwy pliku? [1 - tak, 0 - nie] ({title_from_text}): ").strip() or '1'
        pd_set_input = input(f"Czy dane zużycia liniowego z peakiem na początku: ucinać (1), odwracać (2), nic nie robić (3)? ({pd_set}): ").strip() or '2'
        min_sample_input = input(f"Minimalna ilość danych dla filtru Savitzky-Golay ({min_sample}): ").strip() or str(min_sample)
        max_sample_input = input(f"Maksymalna ilość danych dla filtru Savitzky-Golay ({max_sample}): ").strip() or str(max_sample)
        window_length_u_input = input(f"Długość okna filtru Savitzky-Golay dla µ ({default_window_length_u}): ").strip() or str(default_window_length_u)
        window_length_pd_input = input(f"Długość okna filtru Savitzky-Golay dla pd ({default_window_length_pd}): ").strip() or str(default_window_length_pd)
        chart_lang_input = input(f"Nazwa osi na wykresach [en - angielski, pl - polski] ({chart_lang}): ").strip() or str(chart_lang)
        try:
            offset_raw = int(offset_raw_input)
            title_from_text = int(title_from_text_input)
            pd_set = int(pd_set_input)
            min_sample = int(min_sample_input)
            max_sample = int(max_sample_input)
            window_length_u = int(window_length_u_input)
            window_length_pd = int(window_length_pd_input)
            chart_lang = chart_lang_input
            if window_length_u > int(max_sample_input) or window_length_pd > int(max_sample_input):
                print(f"\033[38;5;214m Liczba większa niż {int(max_sample_input)} \033[0m")
                continue # Powrót na początek pętli
            default_window_length_u = window_length_u
            default_window_length_pd = window_length_pd
            break
        except ValueError:
            print(f"\033[91m Błąd: Wprowadź poprawną liczbę mniejszą niż {int(max_sample_input)}, lub znak. \033[0m")
    if pd_set == 1: erase_peak, invert_peak = 1, 0
    if pd_set == 2: erase_peak, invert_peak = 0, 1
    if pd_set == 3: erase_peak, invert_peak = 0, 0
    print("")
    return min_sample, max_sample, default_window_length_u, default_window_length_pd, title_from_text, offset_raw, erase_peak, invert_peak, chart_lang

# Apply Savitzky–Golay filter to 'µ' and 'pd', leave intact first averaged data value
def Savitzky(averaged_data, column_name, default_window_length):
    """
    Funkcja do zastosowania filtra Savitzky-Golaya na jednej kolumnie.

    Parameters:
    - averaged_data: DataFrame, dane wejściowe.
    - column_name: str, nazwa kolumny do przefiltrowania.
    - default_window_length: int, domyślna długość okna dla filtra.

    Returns:
    - DataFrame z zastosowanym filtrem dla wskazanej kolumny.
    """
    if column_name not in averaged_data:
        raise ValueError(f"[ERROR] Kolumna '{column_name}' nie istnieje w danych.")
    # Oblicz maksymalną dopuszczalną długość okna
    max_window_length = (len(averaged_data) // 2) * 2 - 1  # Upewnij się, że jest nieparzysta
    # Użyj minimalnej z podanej i maksymalnej długości
    window_length = min(default_window_length, max_window_length)
    # Upewnij się, że długość okna jest nieparzysta
    if window_length % 2 == 0:
        window_length += 1
    # Zachowaj pierwszą wartość kolumny
    first_value = averaged_data.at[0, column_name]
    # Zastosuj filtr
    averaged_data[column_name] = savgol_filter(averaged_data[column_name], window_length, polyorder=2)
    # Przywróć pierwszą wartość
    averaged_data.at[0, column_name] = first_value

    return averaged_data

# Funkcja oblicza Srednią dla danej kolumny danych
def display_average(df, column_name):
    """
    Funkcja oblicza i wyświetla średnią wartość z podanej kolumny DataFrame.

    :param df: DataFrame, z którego będą brane dane
    :param nazwa_kolumny: Nazwa kolumny, z której ma być obliczana średnia
    """
    if column_name in df.columns:
        column_mean = df[column_name].mean()
        return column_mean
    else:
        print(f"Kolumna '{column_name}' nie istnieje w DataFrame.")
        return 0

# Funkcja konwertuje przebieg 'µ' z pseudo-sinusoidalnego / prostokątnego na liniowy
# Za pomocą 'Linear Position [mm]' wyznacza odcinki ruchu posuwisto-zwrotnego (linear) [opcjonalnie]
# Tam gdzie była zmiana kierunku zmienia znak i usuwa próbki gdzie ruch ustał lub rósł dopiero
# Po usunięciu aproksymuje te próbki bazując na średniej z dwóch danych przed i po usunięciu
def linear_mode_u_preprocessing(df):
    # Ustawienia
    default_sampling_percent = 0.004  # Domyślnie 0.4% próbek
    max_rows_to_cut = 0  # Suma wierszy wyciętych
    # Sprawdzenie obecności kolumny 'Linear Position [mm]'
    support_column = 'Linear Position [mm]'
    # String START
    lm_warning = "UWAGA! Za mało próbek do ładnego wygładzenia - tylko ABS(), za szybko? za mały promień? za krótko?"
    lm_count = "Ilość zmian znaku:"
    lm_cut = "Liczba wyciętych próbek:"
    lm_per_sample = "próbek na zmianę znaku:"
    # String END
    if support_column in df.columns:
        linear_positions = df[support_column]
        changes_linear = (linear_positions.shift() * linear_positions < 0).cumsum()
        # Zabezpieczenie fluktuacji i obliczanie przesunięcia w 'µ'
        changes_µ = []
        for idx in df.index:
            if idx == 0:
                continue
            if linear_positions.shift()[idx] * linear_positions[idx] < 0:
                # Synchronizacja zmiany znaku w 'µ'
                sign_change_found = False
                rows_after_linear_change = 0
                for i in range(idx, len(df)):
                    if df['µ'].shift()[i] * df['µ'][i] < 0:
                        sign_change_found = True
                        rows_after_linear_change = i - idx
                        break
                if sign_change_found:
                    changes_µ.append((idx, rows_after_linear_change))
        if (round(len(df)/len(changes_µ)) < 3) or (round(len(df)/changes_linear.iloc[-1]) < 3):
            df['µ'] = df['µ'].abs() # ABS wartości w kolumnie 'µ'
            print(lm_warning)
        else:
            # Wycinanie danych na podstawie obliczonego przesunięcia
            for change_idx, delay in changes_µ:
                cut_range = delay # Zakres wycięcia w obie strony
                start_idx = max(0, change_idx - cut_range)
                end_idx = min(len(df) - 1, change_idx + cut_range)
                df.loc[start_idx:end_idx, 'µ'] = None
                max_rows_to_cut += (end_idx - start_idx + 1)
            # Upewnij się, że w zerowej linii 'µ' jest wartość 0
            if df.loc[0, 'µ'] != 0:
                # Przesuń wartości w kolumnie 'µ' w dół o 1
                df['µ'] = df['µ'].shift(1)
                # Dodaj nową wartość 0 na początek kolumny 'µ'
                df.loc[0, 'µ'] = 0
    else:
        # Brak 'Linear Position [mm]' - obliczanie zmian znaku w 'µ'
        changes_µ = changes_linear = (df['µ'].shift() * df['µ'] < 0).cumsum()
        if (round(len(df)/changes_µ.iloc[-1]) < 3):
            df['µ'] = df['µ'].abs() # ABS wartości w kolumnie 'µ'
            print(lm_warning)
        else:
            # Wycinanie danych wokół zmian znaku
            for change_idx in changes_µ.index[df['µ'].shift() * df['µ'] < 0]:
                num_samples = int(len(df) * default_sampling_percent)
                start_idx = max(0, change_idx - num_samples)
                end_idx = min(len(df) - 1, change_idx + num_samples)
                df.loc[start_idx:end_idx, 'µ'] = None
                max_rows_to_cut += (end_idx - start_idx + 1)
    if not ((round(len(df) / changes_linear.iloc[-1]) < 3) if support_column in df.columns else (round(len(df) / changes_µ.iloc[-1]) < 3)):
        # ABS wartości w kolumnie 'µ'
        df['µ'] = df['µ'].abs()
        # Aproksymacja liniowa z użyciem średnich próbek przed i po
        for idx, row in df.iterrows():
            if pd.isna(row['µ']):
                # Znajdź przedział do aproksymacji
                prev_idx = df.loc[:idx - 1, 'µ'].last_valid_index()
                next_idx = df.loc[idx + 1:, 'µ'].first_valid_index()
                if prev_idx is not None and next_idx is not None:
                    # Oblicz wartości średnie przed i po wycięciu
                    pre_values = df.loc[max(0, prev_idx - 1):prev_idx, 'µ'].mean()
                    post_values = df.loc[next_idx:min(next_idx + 1, len(df) - 1), 'µ'].mean()
                    # Aproksymacja liniowa na podstawie średnich wartości
                    num_points = next_idx - prev_idx
                    df.loc[prev_idx:next_idx, 'µ'] = pd.Series(
                        [pre_values + (post_values - pre_values) * (i / num_points)
                        for i in range(num_points + 1)],
                        index=range(prev_idx, next_idx + 1)
                    )
    if max_rows_to_cut > 0:
        if support_column in df.columns:
            print(f"{lm_count} {len(changes_µ)}, {changes_linear.iloc[-1]}")
            if (round(len(df)/len(changes_µ)) > 3):
                print(f"{lm_cut} {max_rows_to_cut}, {lm_per_sample} {round(max_rows_to_cut/len(changes_µ))}")
        else:
            print(f"{lm_count} {changes_µ.iloc[-1]}")
            if (round(len(df)/changes_µ.iloc[-1]) > 3):
                print(f"{lm_cut} {max_rows_to_cut}, {lm_per_sample} {round(max_rows_to_cut/changes_µ.iloc[-1])}")

    return df

# Funkcja usuwa powtarzające się liczby więcej razy niż 10% wszystkich danych
def replace_repeated_values(column, percent=0.1):
    """
    Funkcja usuwa wartości w kolumnie, które powtarzają się więcej niż 10% 
    wszystkich danych. Każda taka wartość jest zastępowana poprzednią "dobrą" 
    wartością, która nie powtarza się ponad próg. Na końcu funkcja zwraca 
    zmodyfikowaną kolumnę oraz wypisuje statystyki.

    Args:
        column (pd.Series): Kolumna danych wejściowych (np. df['µ']).
        percent (float): procent powtarzania, default 0.1, od 0 do 1.

    Returns:
        pd.Series: Zmodyfikowana kolumna z zastąpionymi wartościami.
    """
    # Oblicz liczbę wystąpień każdej wartości
    value_counts = column.value_counts()
    # Oblicz próg powtórzeń (10% wszystkich danych)
    threshold = len(column) * percent
    # Znajdź wartości, które powtarzają się więcej niż 10%
    repeated_values = value_counts[value_counts > threshold].index
   # Jeśli nie ma wartości powtarzających się, zwróć kolumnę bez zmian
    if len(repeated_values) == 0:
        return column
    # Tworzenie kopii kolumny dla przechowania zmodyfikowanych danych
    updated_column = column.copy()
    # Zmienna do przechowywania liczby zamian
    replacement_count = 0
    # Iteracja po wartościach powtarzających się
    for value in repeated_values:
        # Znajdź indeksy wystąpień tej wartości
        indices = column.index[column == value]
        for idx in indices:
            # Znajdź poprzednią dobrą wartość w kolumnie (nie należącą do repeated_values)
            previous_value = None
            for prev_idx in range(idx - 1, -1, -1):  # Iteracja wstecz
                if column.iloc[prev_idx] not in repeated_values:
                    previous_value = column.iloc[prev_idx]
                    break
            # Jeśli znaleziono dobrą wartość, zastąp powtarzającą się
            if previous_value is not None:
                updated_column.at[idx] = previous_value
                replacement_count += 1

    for value in repeated_values:
        print(f"Wartość: {value}, Liczba wystąpień: {value_counts[value]}")
    print(f"UWAGA! Liczba zastąpionych, powtarzających się wartości: {replacement_count}")

    return updated_column

# Srednia ucinana, usuwa wstępnie peaki
def replace_outliers(column, max_mean_range):
    """
    Funkcja zastępuje wartości przekraczające średnią plus/minus max_mean_range 
    średnią wartością danych bez tych pików. Na końcu wyświetla liczbę pików, jeśli są.

    Args:
        column (pd.Series): Kolumna danych wejściowych (np. df['µ']).
        max_mean_range (float): Maksymalny zakres odchylenia od średniej dla wykrywania pików.

    Returns:
        pd.Series: Zmodyfikowana kolumna z zastąpionymi wartościami.
    """
    # Oblicz średnią wartości w kolumnie
    mean = column.mean()
    # Wyznacz granicę dla wykrywania pików
    if mean >= 0:
        max_cut_out = mean + max_mean_range
    else:
        max_cut_out = abs(mean) + max_mean_range
    # Zlicz pikowe wartości
    peak_cnt = column[(column > max_cut_out) | (column < -max_cut_out)].shape[0]
    # Oblicz średnią bez pików
    mean_value = column[(column <= max_cut_out) & (column >= -max_cut_out)].mean()
    # Zastąpienie pików średnią
    updated_column = column.apply(lambda x: mean_value if x > max_cut_out or x < -max_cut_out else x)
    # Wyświetlenie liczby pików, jeśli są
    if peak_cnt > 0:
        print(f"Liczba przekroczeń średniej wartości w kolumnie '{column.name}' ({mean} + {max_mean_range}): {peak_cnt}")

    return updated_column

# Usuwanie peaków bazując na odchyleniu standardowym - W domyśle dla Rtec
def remove_peaks_auto_limit(column, std_multiplier):
    """
    Usuwa duże peaki z podanej kolumny, automatycznie ustalając limit na podstawie średniej i odchylenia standardowego.
    Peaki zastępowane są ostatnią dobrą wartością.

    :param column: Kolumna (Series) do przetworzenia.
    :param std_multiplier: Mnożnik odchylenia standardowego do ustalenia limitu.
    :return: Przetworzona kolumna z usuniętymi peakami.
    """
    count = [0]  # Licznik wykrytych peaków
    # Obliczanie limitów na podstawie średniej i odchylenia standardowego
    mean = column.mean()
    std = column.std()
    lower_limit = mean - std_multiplier * std
    upper_limit = mean + std_multiplier * std
    # Zastępowanie peaków
    def replace_peak(value, prev=[None]):
        if prev[-1] is None:  # Inicjalizacja pierwszej wartości
            prev[-1] = value
        if value < lower_limit or value > upper_limit:  # Jeśli wartość jest poza limitem
            count[0] += 1  # Zwiększ licznik
            return prev[-1]  # Zwróć ostatnią dobrą wartość
        prev[-1] = value  # Zaktualizuj ostatnią dobrą wartość
        return value
    # Przetwarzanie kolumny
    processed_column = column.apply(replace_peak)
    # Wyświetlenie informacji o liczbie wykrytych peaków
    if count[0] > 0:
        print(f"Wykryte i zastąpione peaki w kolumnie '{column.name}': {count[0]} (limit min-max: [{lower_limit:.2f}, {upper_limit:.2f}])")
    
    return processed_column

# Sortowanie df rosnąco
def sort_dataframe_by_column(df, column_name):
    """
    Sortuje DataFrame według podanej kolumny w porządku rosnącym.

    :param df: DataFrame do posortowania.
    :param column_name: Nazwa kolumny, według której ma być posortowany DataFrame.
    :return: Posortowany DataFrame.
    """
    sorted_df = df.sort_values(by=column_name).reset_index(drop=True)
    return sorted_df

# Usuwanie danych poza zakresem
def remove_out_of_range_and_file_limit(df, distance_column, file_name, limit_in):
    """
    Usuwa linie, w których:
    - Wartość w kolumnie distance przekracza limit podany w nazwie pliku.
    - Jeśli w nazwie pliku nie ma liczby, stosuje limit ze zmiennej limit_in.
    
    :param df: DataFrame do przetworzenia.
    :param distance_column: Nazwa kolumny z wartościami odległości.
    :param file_name: Nazwa pliku, z której sprawdzamy limity.
    :return: Przetworzony DataFrame z usuniętymi błędnymi liniami.
    """
    # Etap 1: Poszukiwanie liczby w nazwie pliku w formacie '1000m' lub '1000M' z opcjonalnymi spacjami
    limit_pattern = r'(\d+)\s*[mM]\b'  # Zmienione wyrażenie regularne
    cleaned_file_name = file_name.strip()  # Usuwanie spacji wokół nazwy pliku
    match = re.search(limit_pattern, cleaned_file_name)
    # Ustalenie limitu - jeśli nie znaleziono liczby, ustawiamy domyślny limit 1000
    if match:
        limit = int(match.group(1))
    else:
        limit = limit_in + 1
    # Etap 2: Sprawdzenie, czy wartość w kolumnie distance przekracza limit
    valid_distance_mask = df[distance_column] <= limit
    # Liczba usuniętych wierszy
    removed_count = (~valid_distance_mask).sum()    
    # Usunięcie błędnych wierszy
    cleaned_df = df[valid_distance_mask].reset_index(drop=True)
    # Informacja o liczbie usuniętych wierszy
    if removed_count > 0:
        if match:
            print(f"Usunięto {removed_count} wierszy, które przekroczyły zakres odczytany z pliku {limit}m.")
        else:
            print(f"Usunięto {removed_count} wierszy, które przekroczyły DOMYŚLNY zkres z kodu: {limit_in}m.")
    
    return cleaned_df

# Offset uniwersalny, jeśli jest wartość df poniżej 0
def adjust_negative_offset(df, column):
    """
    Funkcja sprawdza, czy najniższa wartość w danej kolumnie jest mniejsza od zera.
    Jeśli tak, przesuwa wszystkie wartości o wartość offsetu i drukuje informację.
    
    Args:
        df (pd.DataFrame): DataFrame zawierający dane.
        column (str): Nazwa kolumny do sprawdzenia.
    
    Returns:
        pd.DataFrame: DataFrame z poprawionymi wartościami.
    """
    # Sprawdź, czy kolumna istnieje w DataFrame
    if column in df.columns:
        # Znajdź najniższą wartość w kolumnie
        min_value = df[column].min()
        # Jeśli najniższa wartość jest mniejsza od zera, zrób offset
        if min_value < 0:
            print(f"UWAGA! Najniższa wartość w kolumnie '{column}' to {min_value:.4f}. Wykonano ABS().")
            # Znajdź indeks(y) najmniejszej wartości
            min_index = df[df[column] == min_value].index
            # Zmień wartość na jej wartość bezwzględną
            df.loc[min_index, column] = abs(min_value) # abs tylko jednaj wartosci
            #df[column] = df[column] - min_value # offset wszystko o najmniejsza wartosc
            # TODO Poprawić funkcję, inny pomysł ...
    else:
        print(f"Kolumna '{column}' nie istnieje w DataFrame.")
    
    return df

# Wczytanie dla tribometru T11 prędkości liniowej i obciążenia z nazwy pliku z danymi, obliczenie µ i Distance [m]
def T11_calculations(df, file_name):
    s = 0.1 # Default
    F = 10 # Default
    # Dodanie kolumn µ i Distance [m]
    df["µ"] = 0.0  # Placeholder na współczynnik tarcia, do obliczenia poniżej
    df["Distance [m]"] = 0.0  # Placeholder na drogę, do obliczenia poniżej

    # Obliczenie prędkości liniowej i drogi (Distance [m])
    try:
        # Szukanie wartości s z nazwy pliku
        match = re.search(r'(?<!\d)(\d+(?:[.,]\d+)?)(?=\s?(?:m-s|ms|m\s?-\s?s))', file_name) # Obsługa liczb z kropką i przecinkiem
        if match:
            s = float(match.group(1).replace(',', '.'))
            print(f"[T11] Odczytana prędkość liniowa: {s} m/s")
        else:
            # TODO Podanie ręcznie jak nie odczyta z nazwy pliku
            raise ValueError("Brak odpowiedniej wartości m/s w nazwie pliku, np. 0.1m-s [zastosowano domyślnie 0.1m/s]")
        # Obliczanie całkowitej przebytej drogi (m) przy stałej prędkości z każdej chwili czasowej
        df['Distance [m]'] = s * df['Time [s]']
    except Exception as e:
        raise ValueError(f"Problem z obliczeniem drogi: {e}")

    # Obliczenie współczynnika tarcia µ
    try:
        # Szukanie wartości F z nazwy pliku
        match = re.search(r'(?<!\d)(\d+(?:[.,]\d+)?)(?=\s?[Nn])', file_name) # Obsługa liczb z kropką i przecinkiem
        if match:
            F = float(match.group(1).replace(',', '.'))
            print(f"[T11] Odczytana wartość obciążenia: {F} N")
        else:
            # TODO Podanie ręcznie jak nie odczyta z nazwy pliku
            raise ValueError("Brak odpowiedniej wartości F w nazwie pliku, np. 10N [zastosowano domyślnie 10N]")
        # Obliczanie µ = N/F, gdzie N to "Friction force [N]", F odczytane z pliku
        df["µ"] = df["Friction force [N]"].apply(lambda N: N / F if N != 0 else 0)
    except Exception as e:
        raise ValueError(f"Problem z obliczeniem współczynnika tarcia µ: {e}")
    return df

# Wczytanie dla tribometru Rtec prędkości liniowej z nazwy pliku z danymi, konwersja Penetration Depth [µm] i Distance [m]
def Rtec_calculations(df, file_name):
    s = 0.1 # Default
    # Dodanie kolumn Penetration Depth [µm] i Distance [m]
    df["Penetration Depth [µm]"] = 0.0  # Placeholder na zużycie liniowe, do obliczenia poniżej
    df["Distance [m]"] = 0.0  # Placeholder na drogę, do obliczenia poniżej

    # Obliczenie zuźycia liniowego i drogi (Distance [m])
    try:
        # Szukanie wartości s z nazwy pliku
        match = re.search(r'(?<!\d)(\d+(?:[.,]\d+)?)(?=\s?(?:m-s|ms|m\s?-\s?s))', file_name)# Obsługa liczb z kropką i przecinkiem
        if match:
            s = float(match.group(1).replace(',', '.'))
            print(f"[Rtec] Odczytana prędkość liniowa: {s} m/s")
        else:
            # TODO Podanie ręcznie jak nie odczyta z nazwy pliku
            raise ValueError("Brak odpowiedniej wartości m/s w nazwie pliku, np. 0.1m-s [zastosowano domyślnie 0.1m/s]")
        # Obliczanie całkowitej przebytej drogi (m) przy stałej prędkości z każdej chwili czasowej
        df['Distance [m]'] = s * df[' Timestamp']
    except Exception as e:
        raise ValueError(f"Problem z obliczeniem drogi: {e}")

    # Obliczenie Penetration Depth [µm]
    try:        
        # Obliczanie Penetration Depth [µm] z XYZ.Z Position (mm) poprzez pomnożenie przez 1000 aby z mm zrobić µm
        df["Penetration Depth [µm]"] = df["XYZ.Z Position (mm)"] * 1000.0
    except Exception as e:
        raise ValueError(f"Problem z obliczeniem zużycia liniowego: {e}")
    return df

# Główna funkcja wczytująca pliki i dane do DataFrame (df)
# 1. Otwiera plik w trybie odczytu, z kodowaniem UTF-8 i możliwością zastępowania błędów, 
# 2. Wczytuje wszystkie linie z pliku do listy, usuwa znak BOM, 
# 3. Definiowanie typów tribometrów i rodzajów trybu pracy, 
# 4. Wyszukiwanie nagłówka i przypisanie tribometru i rodzaju po nagłówku, 
# 5. Walidacja danych, usunięcie nieprawidłowych, utworzenie DataFrame,
# 6. Podział danych na tribometry i obróbka według każdego rodzaju tribometru i typu
def read_and_process_file(file_path):
    try:
        # Otwórz pliki
        with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
            lines = file.readlines()

        # Usuń BOM z pierwszej linii, jeśli jest obecny (tylko w UTF-8)
        if lines[0].startswith('\ufeff'):
            lines[0] = lines[0][1:]
        
        # Zdefiniuj możliwe typy tribometrów i trybów
        # tribometer_types = ["Nano", "T11", "TRB3", "Rtec"]
        tribometer_types = ["\033[38;5;214mNano\033[0m", "\033[38;5;214mT11\033[0m", "\033[38;5;214mTRB3\033[0m", "\033[38;5;214mRtec\033[0m"] # Nano NTR, T11, TRB3, Rtec
        modes = ["\033[38;5;208mLinear\033[0m", "\033[38;5;208mRotary\033[0m"] # Linear, Rotary

        # Znajdź linię z nagłówkiem i ustal początek danych (ich indeks)
        start_index = None
        tribometer_type = ""
        mode = ""
        
        # Początek danych, typ tribometru i tryb
        for i, line in enumerate(lines):
            # Tryb
            if "Linear mode" in line:
                mode = modes[0]  # "Linear"
            elif "Single-way mode" in line:
                mode = modes[1]  # "Rotary"
            # Typ
            if "Nano Tribometer" in line:
                tribometer_type = tribometer_types[0]  # "Nano" (NTR)
            elif "TRB3" in line:
                tribometer_type = tribometer_types[2]  # "TRB3"
            # Tryb i typ i index
            # Nano Tribometer (NTR)
            if "Time [s]\tDistance [m]\tlaps\tSequence ID\tCycle ID\tMax linear speed [m/s]\tNominal Load [mN]\tµ\tAngle [°]\tNormal force [mN]\tFriction force [mN]\tPenetration depth [µm]" in line:
                mode = modes[1]  # "Rotary" - bo jest Angle [°] w headerze
                tribometer_type = tribometer_types[0]  # "Nano" - bo jest jednostka "mN" w headerze
                start_index = i # i czyli zaczyna od headera
            elif "Time [s]\tDistance [m]\tlaps\tSequence ID\tCycle ID\tMax linear speed [m/s]\tNominal Load [mN]\tµ\tLinear Position [mm]\tNormal force [mN]\tFriction force [mN]\tPenetration depth [µm]" in line:
                mode = modes[0]  # "Linear" - bo jest Linear Position [mm] w headerze
                tribometer_type = tribometer_types[0]  # "Nano" - bo jest jednostka "mN" w headerze
                start_index = i # i czyli zaczyna od headera
            # TRB3
            elif "Time [s]\tDistance [m]\tLaps\tSequence ID\tCycle ID\tMax Linear Speed [m/s]\tNominal Load [N]\tµ\tAngle [°]\tFriction Force [N]\tTemperature [°C]\tHumidity [%]\tPenetration Depth [µm]" in line:
                mode = modes[1]  # "Rotary" - bo jest Angle [°] w headerze
                tribometer_type = tribometer_types[2]  # "TRB3" - bo jest jednostka "N" oraz Temperature [°C] i Humidity [%] w headerze
                start_index = i # i czyli zaczyna od headera
            elif "Time [s]\tDistance [m]\tLaps\tSequence ID\tCycle ID\tMax Linear Speed [m/s]\tNominal Load [N]\tµ\tLinear Position [mm]\tFriction Force [N]\tTemperature [°C]\tHumidity [%]\tPenetration Depth [µm]" in line:
                mode = modes[0]  # "Linear" - bo jest Linear Position [mm] w headerze
                tribometer_type = tribometer_types[2]  # "TRB3" - bo jest jednostka "N" oraz Temperature [°C] i Humidity [%] w headerze
                start_index = i # i czyli zaczyna od headera
            # T11
            elif "Time [s];Friction force [N];Displacement [um];Temperature2 [C];Temperature1 [C];Rotational speed [rpm];Number of revolutions" in line:
                mode = modes[1]  # "Rotary" - bo jest Number of revolutions w headerze
                tribometer_type = tribometer_types[1]  # "T11" - bo separator ";" oraz Temperature1 [C] i Temperature2 [C]
                start_index = i # i czyli zaczyna od headera
            # Rtec
            elif "Step, Timestamp, RecipeStep, DAQ.Fz (N),DAQ.Fx (N),DAQ.COF (),Rotary.Velocity (rpm),XYZ.Z Depth (mm),XYZ.Z Position (mm),Rotary.Angle (deg)," in line:
                mode = modes[1]  # "Rotary" - bo jest Rotary.Angle (deg) w headerze
                tribometer_type = tribometer_types[3]  # "Rtec" - bo jest separator "," oraz inne nagłówki i jednostki w "()"
                start_index = i # i czyli zaczyna od headera
        
        if start_index is None:
            raise ValueError(f"\033[91m Nie znaleziono odpowiedniej linii rozpoczynającej dane w pliku: {file_path} \033[0m")
        
        # Wczytaj nagłówek i dane
        # HEADER
        # .rstrip() - Usuń znaki końca linii i separator na końcu jak są
        # Dla T11 ";" jako separator
        if tribometer_type == tribometer_types[1]:  # "T11"
            header = lines[start_index].strip().rstrip(';\r\n').split(';')
        # Dla TRB3 i Nano TRB to TAB (\t) jako separator
        elif tribometer_type == tribometer_types[2] or tribometer_type == tribometer_types[0]:  # "TRB3 i Nano"
            header = lines[start_index].strip().rstrip('\t\r\n').split('\t')
        # Dla Rtec "," jako separator
        elif tribometer_type == tribometer_types[3]:  # "Rtec"
            header = lines[start_index].strip().rstrip(',\r\n').split(',')
        # W innym przypadku lub błędu
        else:
            raise ValueError(f"\033[91m Nie znaleziono separatora nagłówków \033[0m")

        # DANE
        # Dla T11 i Rtec dodatkowy poniżej fix na usunięcie znaku separatora na końcu linii i usunięcie linii z wartościami zerowymi w kolumnie "Friction force [N]"
        data = []
        valid_rows = 0
        invalid_rows = 0
        for line in lines[start_index + 1:]:
            # Dla T11 ";" jako separator
            if tribometer_type == tribometer_types[1]:  # "T11"
                line = line.strip().rstrip('\r\n;') # Usuń znaki końca linii i separator na końcu jak są
                row = line.strip().split(';')
            # Dla Rtec "," jako separator
            elif tribometer_type == tribometer_types[3]:  # "Rtec"
                line = line.strip().rstrip('\r\n,') # Usuń znaki końca linii i separator na końcu jak są
                row = line.strip().split(',')
            # Dla TRB3 i Nano TRB to TAB (\t) jako separator
            elif tribometer_type == tribometer_types[2] or tribometer_type == tribometer_types[0]:  # "TRB3 i Nano"
                line = line.strip().rstrip('\r\n\t') # Usuń znaki końca linii i separator na końcu jak są
                row = line.strip().split('\t')
            # W innym przypadku lub błędu
            else:
                raise ValueError(f"\033[91m Nie znaleziono separatora danych \033[0m")
            
            # Walidacja, czy wiersz zawiera odpowiednią liczbę kolumn, jak nie to pomiń
            if len(row) != len(header):
                invalid_rows += 1
                continue

            # Walidacja, czy wszystkie elementy można skonwertować na float (zamiana ewentualnie , na .)
            try:
                #row = [float(value.replace(',', '.')) for value in row]
                #data.append(row)
                #liczba_poprawnych_wierszy += 1
                for i in range(len(row)):
                    if ',' in row[i]: # jeśli przecinek
                        row[i] = row[i].replace(',', '.') # konwertuj na kropkę
                    if not re.match(r'^-?\d+(?:\.\d+)?$', row[i]): # jeśli jest tekst to pomijaj linię
                        invalid_rows += 1
                        break
                    else:
                        row[i] = float(row[i]) # jeśli wszystko w porządku, konwertuj na float
                else:
                    data.append(row) # przypisz do nowej macierzy dane wyselekcjonowane / poprawne
                    valid_rows += 1

            except ValueError as e:
                #print(f"\033[91m Błąd konwersji wartości na float w wierszu: {row} \033[0m")
                invalid_rows += 1
                continue
        
        if invalid_rows > 0:
            print(f"[{tribometer_type}] \033[91mLiczba niepoprawnych wierszy: {invalid_rows} z {valid_rows + invalid_rows}\033[0m")

        # Usuń wiersze z brakującymi lub nieprawidłowymi danymi
        data = [row for row in data if all(value != '' and value is not None for value in row)]

        # Sprawdź, czy są jakieś poprawne dane do przetworzenia
        if not data:
            raise ValueError(f"\033[91m Brak poprawnych danych w pliku {file_path} \033[0m")

        # Stwórz DataFrame
        df = pd.DataFrame(data, columns=header)

        # Sprawdź, czy plik zawiera kolumnę "Linear Position [mm]" jak tak to ruch Posuwisto-Zwrotny
        if mode == modes[0]:  # "Linear"
            if tribometer_type == tribometer_types[2] or tribometer_type == tribometer_types[0]:  # "TRB3 i Nano"
                if tribometer_type == tribometer_types[0]:  # "Nano" (NTR) - JUST IN CASE
                    # Set 0 to all rows in this column (just in case)
                    df['Penetration Depth [µm]'] = 0
                    # Obliczenie średniej ucinanej dla wartości bez pików
                    df['µ'] = replace_outliers(df['µ'], 1.0)

                # Usuń powtarzające się te same liczby, więcej niż 10% zbioru danych (FIX)
                df['µ'] = replace_repeated_values(df['µ'])
                # Konwersja przebiegu 'µ' z pseudo-sinusoidalnego / prostokątnego na liniowy
                df = linear_mode_u_preprocessing(df)
                # Usunięcie peaków za pomocą odchylenia standardowego razy 2
                df['µ'] = remove_peaks_auto_limit(df['µ'], std_multiplier=2)
                # Wybierz interesujące kolumny
                df = df[['Distance [m]', 'µ', 'Penetration Depth [µm]']]
            else:
                raise ValueError(f"\033[91m Brak funkcji programu typu: {mode}, dla tego tribometru: {tribometer_type} \033[0m")
        # Jak nie to ruch Obrotowy
        elif mode == modes[1]:  # "Rotary"
            # Jeśli jest to Nano TRB (NTR)
            if tribometer_type == tribometer_types[0]:  # "Nano"
                # Set 0 to all rows in this column (just in case)
                df['Penetration Depth [µm]'] = 0
                # Obliczenie średniej ucinanej dla wartości bez pików
                df['µ'] = replace_outliers(df['µ'], 1.0)
                # Usuń powtarzające się te same liczby, więcej niż 10% zbioru danych (FIX)
                df['µ'] = replace_repeated_values(df['µ'])
                # Wybierz interesujące kolumny
                df = df[['Distance [m]', 'µ', 'Penetration Depth [µm]']]

            # Jeśli to T11
            elif tribometer_type == tribometer_types[1]:  # "T11"
                # Usuń powtarzające się te same liczby, więcej niż 5% zbioru danych (FIX)
                df['Friction force [N]'] = replace_repeated_values(df['Friction force [N]'], 0.05)

                # # Znajdź najniższą wartość w kolumnie "Friction force [N]"
                #df = adjust_negative_offset(df, 'Friction force [N]')

                # Zmień nazwy kolumny
                df = df.rename(columns={'Displacement [um]': 'Penetration Depth [µm]'})
                # Dodanie kolumn "Distance [m]" oraz "µ" i obliczenie tych danych dla Tribometru T11
                df = T11_calculations(df, os.path.basename(file_path))
                
                # Wybierz interesujące kolumny
                df = df[['Distance [m]', 'µ', 'Penetration Depth [µm]']]
            # Jeśli to TRB3
            elif tribometer_type == tribometer_types[2]:  # "TRB3"

                # # Znajdź najniższą wartość w kolumnie "µ"
                #df = adjust_negative_offset(df, 'µ')

                # Wybierz interesujące kolumny
                df = df[['Distance [m]', 'µ', 'Penetration Depth [µm]']]
            # Jeśli to Rtec
            elif tribometer_type == tribometer_types[3]:  # "Rtec"
                # Usuń powtarzające się te same liczby, więcej niż 5% zbioru danych (FIX)
                df['DAQ.COF ()'] = replace_repeated_values(df['DAQ.COF ()'], 0.05)

                # Obliczenie średniej ucinanej dla wartości bez pików - wymaga bo Rtec ma dużo pików --- µ ---
                #df['DAQ.COF ()'] = replace_outliers(df['DAQ.COF ()'], 0.5)

                # Zmień nazwy kolumny
                df = df.rename(columns={'DAQ.COF ()': 'µ'})

                # Konwersja XYZ.Z Position (mm) na um oraz obliczenie Distance [m] z Time [s]
                df = Rtec_calculations(df, os.path.basename(file_path))

                # Sortowanie rosnąco według 'Distance [m]' (usuwa dziwne anomalie w danych)
                df = sort_dataframe_by_column(df, 'Distance [m]')
                
                # usuwanie danych poza zakresem - 1000 m maksymalny zakres jak nie będzie podany w nazwie pliku
                df = remove_out_of_range_and_file_limit(df, 'Distance [m]', os.path.basename(file_path), 1000)

                # Usunięcie peaków za pomocą odchylenia standardowego razy 3 --- µ ---
                df['µ'] = remove_peaks_auto_limit(df['µ'], std_multiplier=3)

                # Usunięcie peaków za pomocą odchylenia standardowego razy 2 --- Penetration Depth [µm] ---
                df['Penetration Depth [µm]'] = remove_peaks_auto_limit(df['Penetration Depth [µm]'], std_multiplier=2)

                # Wybierz interesujące kolumny
                df = df[['Distance [m]', 'µ', 'Penetration Depth [µm]']]
            else:
                raise ValueError(f"\033[91m Brak funkcji programu typu [{mode}] dla tego tribometru: {tribometer_type} \033[0m")
        # Jak nie Obrotowy czy Posuwisto-zwrotny
        else:
            raise ValueError(f"\033[91m Brak wykrytego poprawnie rodzaju ruchu lub nie zdefiniowany, dla: {file_path} \033[0m")
        return df, tribometer_type, mode

    except Exception as e:
        print(f"\033[91m Błąd podczas przetwarzania pliku: {file_path}: {e} \033[0m")
        return None, None, None

# Funkcja oblicza najlepszą wartość ilości uśredniania próbek dla przebiegu µ i pd
#column_names = ['Distance [m]', 'µ', 'Penetration Depth [µm]']
def find_optimal_samples_average(data, column_names, min_sample, max_sample):
    min_sample_average = max(1, len(data) // max_sample)  # Wyznacz dolną granicę `sample_average` (MIN)
    # Dla max_sample = 500, jeśli len(data) = 1000, to min_sample_average = 2.
    max_sample_average = max(1, len(data) // min_sample)  # Wyznacz górną granicę `sample_average` (MAX)
    # Dla min_sample = 100, jeśli len(data) = 1000, to max_sample_average = 10.
    
    best_sample_average_µ = None
    best_sample_average_pd = None
    best_std_µ = float('inf')
    best_std_pd = float('inf')

    # Przechodzimy przez różne wartości `sample_average`
    for sample_average in range(min_sample_average, max_sample_average + 1):
        averaged_data_list = []

        # Tworzymy uśrednione dane
        for i in range(0, len(data), sample_average):
            subset = data.iloc[i:i+sample_average]
            
            if len(subset) < 2:
                continue
            
            µ_avg = subset['µ'].mean()
            pd_avg = subset['Penetration Depth [µm]'].mean()

            averaged_data_list.append({
                'Distance [m]': subset.iloc[0]['Distance [m]'],
                'µ': µ_avg,
                'Penetration Depth [µm]': pd_avg
            })
        
        averaged_data = pd.DataFrame(averaged_data_list, columns=column_names)
        std_µ = averaged_data['µ'].std()
        std_pd = averaged_data['Penetration Depth [µm]'].std()

        # Aktualizujemy najlepsze wartości `sample_average`
        if std_µ < best_std_µ:
            best_std_µ = std_µ
            best_sample_average_µ = sample_average
            
        if std_pd < best_std_pd:
            best_std_pd = std_pd
            best_sample_average_pd = sample_average
    
    return best_sample_average_µ, best_sample_average_pd  # Zwracamy krotkę (best_sample_average_µ, best_sample_average_pd)

# Funkcja dokonuje korekty wykresu zużycia liniowego aby zaczynał się od zera
# 1. spr. czy jest kolumna z danymi, 2. offset o najmniejszą wartość danych, 3. offset o najmniejsze lokalne minimum dla dodatnych danych, 
# 4. wykrycie piku ujemnego i go usunięcie z początka danych, 5. jak nie ma piku to offset o najniższą wartość w danych, 
# 6. Offset danych RAW aby przyrównać z danymi uśrednionymi, 7. usunięcie dużego peaku z początku danych uśrednionych, 
# 8. Inwersja dużego peaku z początku danych i offset w górę 9. wstaw 0 na początek danych, jak nie ma.
# df - dane uśrednione, data - dane RAW (czyste), percent - proc. danych do analizy piku, offset_raw - przyrównanie data do df, 
# erase_peak - usuwa znaczący pik danych, invert_peak - inwersja i dodanie offsetu piku danych
def process_penetration_depth(df, data, percent, offset_raw, erase_peak, invert_peak):   
    # Sprawdź, czy kolumna istnieje i czy wszystkie wartości są równe zero, jeśli nie istnieje albo są same 0 to pomiń cały kod i zwróć DataFrame
    if 'Penetration Depth [µm]' not in df.columns or (df['Penetration Depth [µm]'] == 0).all():
        return df  # Zwróć df, jeśli kolumna nie istnieje
    else:
        # DLA WYKRESÓW UJEMNYCH
        # Oblicz i zastosuj offset dla najmniejszej wartości
        min_penetration_depth_1 = df['Penetration Depth [µm]'].min()  # Znajdź najmniejszą wartość
        if min_penetration_depth_1 < 0:  # Sprawdź, czy najmniejsza wartość jest mniejsza od 0
            df['Penetration Depth [µm]'] += abs(min_penetration_depth_1)  # Zastosuj offset do wszystkich wartości
            # ^^- przesunięcie wykresu w górę o najmniejszą wartość zbioru danych

        # DLA WYKRESÓW CAŁYCH DODATNICH (po powyższej korekcie)
        # Oblicz i zastosuj offset dla najmniejszego minimum lokalnego
        local_minima = (df['Penetration Depth [µm]'].shift(1) > df['Penetration Depth [µm]']) & (df['Penetration Depth [µm]'].shift(-1) > df['Penetration Depth [µm]'])
        min_penetration_depth_2 = df['Penetration Depth [µm]'][local_minima].min()  # Znajdź najmniejsze minimum lokalne
        if min_penetration_depth_2 > 0:  # Sprawdź, czy najmniejsze minimum lokalne jest większe od 0
            df['Penetration Depth [µm]'] += (-(min_penetration_depth_2))  # Zastosuj offset do wszystkich wartości
            # ^^- przesunięcie wykresu w dół o najmniejsze dodatnie minimum

        # DLA WYKRESÓW Z PIKIEM UJEMNYM NA POCZĄTKU (usuwanie ujemnych pików)
        # Sprawdzenie, czy pierwsza wartość w kolumnie Penetration Depth [µm] jest mniejsza od 0
        set_offset = False
        if df.iloc[0]['Penetration Depth [µm]'] < 0:
            # Oblicz procent wartości ze wszystkich wierszy (10% jako default dla zmiennej percent)
            limit = int(len(df) * percent)
            for index in range(1, min(limit, len(df))):
                if df.iloc[index]['Penetration Depth [µm]'] > 0:
                    indices_to_zero = list(range(1, index + 1))  # Indeksy human-readable
                    df.iloc[:index, df.columns.get_loc('Penetration Depth [µm]')] = 0  # Ustaw 0 dla wszystkich wartości ujemnych przed tą wartością dodatnią
                    print(f"Usunięto początkowy pik danych 'pd', od {indices_to_zero[0]} do {indices_to_zero[-1]}")
                    if erase_peak == 0 and indices_to_zero[-1] > 1: print(f"UWAGA! Usunięto część danych 'pd', dane błędne??? Porównaj z RAW")
                    break
        # DLA WYKRESÓW UJEMNYCH ALE ROSNĄCYCH (cały czas albo prawie cały czas)
            else:
                # Jeśli wykres jest nadal ujemny powyżej limitu (10% jako default dla zmiennej percent)
                min_penetration_depth_3 = df['Penetration Depth [µm]'].min()  # Oblicz minimum z całej kolumny
                if min_penetration_depth_3 < 0:  # Sprawdź, czy minimum jest mniejsze od 0
                    df['Penetration Depth [µm]'] += abs(min_penetration_depth_3)  # Zastosuj offset, aby najmniejsza wartość była maksymalnie zero
                    set_offset = True
        # Dodatkowy offset wykresów rosnących (operacje już na wartościach dodatnich)
        # [oryginalne dane nie są uśredniane a początkowe wartości mogą się szybko zmieniać, więc poniższy kod to taki fix]
        if set_offset == True:
            # Znajdź minimum w kolumnie
            min_value = df['Penetration Depth [µm]'].min()
            min_index = df['Penetration Depth [µm]'].idxmin()
            # Sprawdź pozycję minimum
            if min_index in [0, 1]:
                # Jeśli minimum jest na początku, użyj różnicy pierwszych dwóch wartości
                difference = abs(df['Penetration Depth [µm]'].iloc[1] - df['Penetration Depth [µm]'].iloc[0])
                df['Penetration Depth [µm]'] += abs(difference)  # Zastosuj offset
                print(f"Zastosowano dodatkowy offset {abs(difference):.1f} bazujący na różnicy pierwszych dwóch wartości.")
                set_offset = False
            elif min_index in [len(df) - 2, len(df) - 1]:
                # Jeśli minimum jest na końcu, użyj różnicy ostatnich dwóch wartości
                difference = abs(df['Penetration Depth [µm]'].iloc[-1] - df['Penetration Depth [µm]'].iloc[-2])
                df['Penetration Depth [µm]'] += abs(difference)  # Zastosuj offset
                print(f"Zastosowano dodatkowy offset {abs(difference):.1f} bazujący na różnicy ostatnich dwóch wartości.")
                set_offset = False

        if erase_peak == 1 and invert_peak == 0: # (tylko dla wykresów powyżej lub równo z 0, czyli po korektach)
            # Usunięcie peaku zgodnie z zaleceniem Prof. MM
            min_value = df['Penetration Depth [µm]'].min()  # Znajdź najmniejszą wartość w kolumnie
            end_index = df[df['Penetration Depth [µm]'] == min_value].index  # Ustal indeks najmniejszej wartości
            if min_value <= 1.0 and not isinstance(end_index, pd.RangeIndex): # Tylko dla wartości mniejszych od 1 i nie dla RangeIndex
                if end_index[0] > 0: # Jeśli index nie jest zerowy (czyli nie pierwsza pozycja)
                    df.loc[:end_index[0], 'Penetration Depth [µm]'] = 0 # Ustaw wartości w kolumnie na 0 od 0 do pozycji end_index[0]
                elif len(end_index) > 1: # Jeśli będzie wiecej indeksów
                    df.loc[:end_index[-1], 'Penetration Depth [µm]'] = 0  # Ustaw wartości w kolumnie na 0 od 0 do pozycji ostatniego end_index

        elif erase_peak == 0 and invert_peak == 1: # (tylko dla wykresów powyżej lub równo z 0, czyli po korektach)
            # dodanie peaku jako inwersja i offset na początku danych
            consecutive_zero_min_index = 0  # Licznik wystąpień min_index = 0
            # Pętla while jest po to aby kod wykonać do puki nie wygładzi się przebiegów pseudo-sinusoidalnych
            while consecutive_zero_min_index < 3:  # Warunek zakończenia pętli - POTRÓJNY check, jeśli indeks się trzy razy nie zmieni to break
                min_value = df['Penetration Depth [µm]'].min()  # Znajdź najmniejszą wartość w kolumnie
                min_index = df['Penetration Depth [µm]'].idxmin()  # Znajdź indeks najmniejszej wartości
                end_index = df[df['Penetration Depth [µm]'] == min_value].index  # Ustal indeks najmniejszej wartości
                if min_value <= 0 and not isinstance(end_index, pd.RangeIndex):  # Tylko dla wartości mniejszych od 0 i nie dla RangeIndex
                    for i in range(len(end_index)):  # iterowanie po całym zakresie indeksów
                        df.loc[:end_index[i], 'Penetration Depth [µm]'] *= -1  # Inwertuj wartości w kolumnie od 0 do pozycji end_index[i]
                        min_value = df['Penetration Depth [µm]'].min()  # Znajdź najmniejszą wartość w kolumnie po inwersji
                        df['Penetration Depth [µm]'] += abs(min_value)  # Offsetuj całe dane, aby najmniejsza wartość była większa bądź równa 0
                if min_index == 0:  # Sprawdź, czy min_index jest równy 0
                    consecutive_zero_min_index += 1  # Zwiększ licznik
                else:
                    consecutive_zero_min_index = 0  # Zresetuj licznik, jeśli min_index nie jest równy 0

        # Drugi raz offset wartości ujemnych (na wszelki wypadek) względem danych uśrednionych
        min_penetration_depth_4 = df['Penetration Depth [µm]'].min()  # Znajdź najmniejszą wartość
        if min_penetration_depth_4 < 0:  # Sprawdź, czy najmniejsza wartość jest mniejsza od 0
            df['Penetration Depth [µm]'] += abs(min_penetration_depth_4)  # Zastosuj offset do wszystkich wartości
            if df.loc[0, 'Penetration Depth [µm]'] != 0: df.loc[0, 'Penetration Depth [µm]'] = 0 # Ustaw jeszcze raz 0

        if offset_raw == 1:
            # Offset danych RAW aby przyrównać z danymi uśrednionymi
            middle_index_df = len(df) // 2
            middle_index_data = len(data) // 2
            range_data = int(len(data) * 0.05)  # 5% zakresu danych
            # Oblicz indeksy dla uśrednionego odcinka danych
            start_index_data = max(0, middle_index_data - range_data)  # Indeks początkowy
            end_index_data = min(len(data), middle_index_data + range_data)  # Indeks końcowy
            # Oblicz średnią wartość dla uśrednionego odcinka danych
            average_penetration_depth = data['Penetration Depth [µm]'].iloc[start_index_data:end_index_data].mean()
            if df['Penetration Depth [µm]'].iloc[middle_index_df] != average_penetration_depth:
                difference = average_penetration_depth - df['Penetration Depth [µm]'].iloc[middle_index_df]
                data['Penetration Depth [µm]'] += -difference  # Dodaj offset danych w Penetration Depth [µm]

        # Upewnij się, że w zerowej linii "Penetration Depth [µm]" jest wartość 0
        if df.loc[0, 'Penetration Depth [µm]'] != 0:
            # Przesuń wartości w kolumnie 'Penetration Depth [µm]' w dół o 1
            df['Penetration Depth [µm]'] = df['Penetration Depth [µm]'].shift(1)
            # Dodaj nową wartość 0 na początek kolumny 'Penetration Depth [µm]'
            df.loc[0, 'Penetration Depth [µm]'] = 0
        else:
            # Przypisz wartościom zerowym ostatnią znaną wartość powyżej zera
            # Znalezienie wartości 0 od indeksu 1 w kolumnie
            min_values = df['Penetration Depth [µm]'].iloc[1:].min()  # Znajdź wartości 0 od indeksu 1
            min_indexes = df[df['Penetration Depth [µm]'] == min_values].index  # Znajdź indeksy wszystkich wartości 0
            for index in min_indexes:
                if index > 0 and min_values == 0:  # Jeśli indeks większy od 0 i wartość minimalna równa 0
                    # Znajdź następną wartość różną od 0
                    next_nonzero_idx = df['Penetration Depth [µm]'].iloc[index + 1:].ne(0).idxmax()  # Indeks następnej wartości != 0
                    if next_nonzero_idx > index:  # Jeśli istnieje następna wartość różna od 0
                        prev_value = df['Penetration Depth [µm]'].iloc[index - 1]  # Wartość poprzednia
                        next_value = df['Penetration Depth [µm]'].iloc[next_nonzero_idx]  # Wartość następna
                        # Aproksymacja liniowa
                        interpolated_value = prev_value + (next_value - prev_value) / (next_nonzero_idx - index)
                        df.loc[index, 'Penetration Depth [µm]'] = interpolated_value  # Przypisanie aproksymowanej wartości
                        print(f"Zinterpolowano wartość 'pd' dla indeksu {index + 1}: {interpolated_value:.1f}")  # Drukuj indeks human-readable
            if erase_peak == 1 and min_indexes > 0: print("UWAGA! Usunięto zera po piku w 'pd', według configu użytkownika programu.") # Drukuj info
        
        return df

# Funkcja uśrednia RAW Data według zmiennej best_sample_average i stasuje filtr Savitzky-Golay
def adjust_and_average_data(df, best_sample_average, window_u, window_pd):
    # Zainicjuj pustą listę na uśrednione dane
    averaged_data_list = []

    # Iteruj po zakresach danych co `sample_average` wierszy
    for i in range(0, len(df), best_sample_average):
        # Wybierz podzbiór danych dla danego zakresu
        subset = df.iloc[i:i+best_sample_average]

        # Oblicz średnią dla µ i Penetration Depth [µm] (dodatkowo zabezpieczenie 1: wartość µ nie może być ujemna)
        µ_avg = subset['µ'].mean()
        # if µ_avg < 0:
        #     µ_avg = abs(µ_avg)
        # TODO dodać obsługę nano tribometru (NTR) - TZN. współczynnik tarcia scgodzi na ujemne wartości, jaki fix? abs? offset?
        pd_avg = subset['Penetration Depth [µm]'].mean()

        # Dodaj uśrednione wartości do listy
        averaged_data_list.append({
            'Distance [m]': subset.iloc[0]['Distance [m]'],  # Zachowaj oryginalną wartość Distance [m]
            'µ': µ_avg,
            'Penetration Depth [µm]': pd_avg
        })

    # Stwórz DataFrame z uśrednionymi danymi
    averaged_data = pd.DataFrame(averaged_data_list)

    # Filtrowanie kolumny 'µ' z długością okna 'window_u'
    averaged_data = Savitzky(averaged_data, 'µ', window_u)

    # Filtrowanie kolumny 'Penetration Depth [µm]' z długością okna 'window_pd'
    averaged_data = Savitzky(averaged_data, 'Penetration Depth [µm]', window_pd)
    
    # Upewnij się, że w zerowej linii "µ" jest wartość 0
    if averaged_data.loc[0, 'µ'] != 0:
        # Przesuń wartości w kolumnie 'µ' w dół o 1
        averaged_data['µ'] = averaged_data['µ'].shift(1)
        # Dodaj nową wartość 0 na początek kolumny 'µ'
        averaged_data.loc[0, 'µ'] = 0

    # Upewnij się, że w zerowej linii "Distance [m]" jest wartość 0
    if averaged_data.loc[0, 'Distance [m]'] != 0:
        # Przesuń wartości w kolumnie 'Distance [m]' w dół o 1
        averaged_data['Distance [m]'] = averaged_data['Distance [m]'].shift(1)
        # Dodaj nową wartość 0 na początek kolumny 'Distance [m]'
        averaged_data.loc[0, 'Distance [m]'] = 0

    "Poniżej jest usuwanie powtórzeń wierszy, usuwanie o małej gęstości pomiarów, zostawienie kilku w przypadku braku danych"
    # Oblicz różnice bezwzględne między kolejnymi wierszami w 'Distance [m]'
    differences = averaged_data['Distance [m]'].diff().abs()
    if differences.any():
        # Usuń wiersze z danych względem kolumny "Distance [m]" tak, aby były maksymalnie z gęstością wartości co max 0.15
        filtered_data = averaged_data[(differences > 0.15) | (differences.isna())]
        # Sprawdź, czy po filtracji jest mniej niż dwa wiersze
        if filtered_data.shape[0] < 2:
            # Zachowaj tylko kilka wierszy z "Distance [m]" < 0.1
            averaged_data = averaged_data[averaged_data['Distance [m]'] < 0.1]
        else:
            # Jeśli jest więcej jak dwa wiersze to przypisz przefiltrowane dane do averaged_data
            averaged_data = filtered_data
    else:
        # Zostaw tylko dwa wiersze dla tych samych wartości Distance [m]
        averaged_data = averaged_data.groupby('Distance [m]').head(1 + 2)

    return averaged_data

# Funkcja dodaje ostatni pomiar z aproksymacji pomiarów, jeżeli brakuje ostatniego wiersza / linii
def approximate_last_measurement(averaged_data, original_data):
    if averaged_data.shape[0] < 2:  # Sprawdza liczbę wierszy
        print("\033[91mBrak wystarczających danych do przybliżenia\033[0m")  # Not enough data to approximate
        return averaged_data
    
    # Check if last two Distance [m] values are the same
    if averaged_data['Distance [m]'].iloc[-1] == averaged_data['Distance [m]'].iloc[-2]:
        return averaged_data  # Skip if last two Distance [m] values are the same

    # TODO Remove last zero or negative values from rows in averaged_data (OPTIONAL, never happend?)

    # Get the original last Distance [m] value
    original_last_distance = original_data['Distance [m]'].iloc[-1]

    # Find the index of original_data where Distance [m] is the original_last_distance
    original_last_index = original_data.index[original_data['Distance [m]'] == original_last_distance]

    # Check if the index is not empty
    if not original_last_index.empty:
        # Get the index location
        original_last_index = original_last_index.values[0]

        # Create a new row with the last averaged values and original last Distance [m]
        new_row = {
            'Distance [m]': original_data.loc[original_last_index, 'Distance [m]'],
            'µ': averaged_data['µ'].iloc[-1],
            'Penetration Depth [µm]': averaged_data['Penetration Depth [µm]'].iloc[-1] if 'Penetration Depth [µm]' in averaged_data else None
        }

        # Concatenate the new row to averaged_data if Distance [m] value is not duplicated
        if new_row['Distance [m]'] not in averaged_data['Distance [m]'].values:
            averaged_data = pd.concat([averaged_data, pd.DataFrame([new_row])], ignore_index=True)

    return averaged_data

# Funkcja generująca wykresy z danych
def generate_combined_xlsx(csv_files, output_xlsx, series_from_filename, chart_lang):
    with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
        col_offset = 0  # Przesunięcie kolumn na dane do wykresu (współrzędne do wykresów)
        row_offset = 0  # Przesunięcie wiersza na wykresy
        sheet_name = output_xlsx.replace('.xlsx', "")[:31]  # Nazwa arkusza danych (limit do 31 znaków)

        # Ustawienie nazw osi w zależności od języka
        if chart_lang == 'pl':
            x_axis_label = 'Droga [m]'
            y_axis_label_mu = 'Współczynnik tarcia'
            y_axis_label_pd = 'Zużycie liniowe [µm]'
        else:  # Domyślnie angielski
            x_axis_label = 'Distance [m]'
            y_axis_label_mu = 'Friction coefficient'
            y_axis_label_pd = 'Linear wear [µm]'

        for csv_file in csv_files:
            # Wczytaj dane z pliku CSV do DataFrame
            df = pd.read_csv(csv_file)  # Pobranie danych z plików .csv do Dataframe

            # Zapisz dane z DataFrame do arkusza, z uwzględnieniem przesunięcia kolumn
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=col_offset)    

            # Wyodrębnij nazwę pliku i tekst w nawiasach
            filename = os.path.basename(csv_file)  # Nazwa pliku
            match = re.search(r'\((.*?)\)', filename)  # szukaj nawiasu i danych w nim
            text_in_brackets = match.group(1) if match else ""  # jak jest nawias to przypisz dane
            cleaned_filename = filename.replace('.csv', "").replace('(', "").replace(')', "")  # usuń nawias z nazwy

            # Uzyskaj obiekt worksheet
            worksheet = writer.sheets[sheet_name]
            # Info na początku pliku .xlsx
            worksheet.write(0, col_offset, cleaned_filename) # nazwa bez nawiasów w pierwszej kolumnie

            # Jeśli są dane w nawiasie, to zapisz je
            series_name = text_in_brackets if match else filename
            # Jeśli chcesz nazwę serii z nazwy pliku w nawiasie, to ją przypisz do serii
            if series_from_filename == 1:
                worksheet.write(1, col_offset + 2, series_name) # nazwa z zawartością nawiasów w trzeciej kolumnie
                worksheet.write(1, col_offset + 1, series_name)

            # Ustal maksymalną wartość dla osi X
            x_axis_max = int(round(df['Distance [m]'].max()))
            x_axis_max = (int((x_axis_max + 24) / 25)) * 25

            # Dodaj wykres µ
            if series_from_filename == 0: worksheet.write(1, col_offset + 1, 'µ')
            chart = writer.book.add_chart({'type': 'scatter'})
            chart.add_series({
                'name': [sheet_name, 1, col_offset + 1],  # Nazwa serii z komórki 'µ'
                'categories': [sheet_name, 3, col_offset, 3 + len(df) - 1, col_offset],  # Zakres dla osi X
                'values': [sheet_name, 3, col_offset + 1, 3 + len(df) - 1, col_offset + 1],  # Zakres dla osi Y
                'line': {'width': 2},  # Ciągła linia bez punktów
                'marker': {'type': 'none'},  # Wyłącza wyświetlanie punktów
            })
            chart.set_title({'name': f"{cleaned_filename} - µ"})
            chart.set_x_axis({
                'name': x_axis_label,  # nazwa osi X zależna od języka
                'min': 0,  # minimalna wartość osi X
                'max': 100 if x_axis_max < 10 else x_axis_max,  # zakres
                'major_unit': max(25, int(round(x_axis_max / 5.0))) if x_axis_max >= 10 else 2,  # podziałka
                'major_gridlines': {'visible': True},  # dodaj pionowe kreski
            })
            chart.set_y_axis({
                'name': y_axis_label_mu,  # nazwa osi Y zależna od języka
            })
            worksheet.insert_chart(f"B{25 + row_offset}", chart)

            # Dodaj wykres Penetration Depth [µm] (jeśli istnieje)
            if 'Penetration Depth [µm]' in df.columns:
                if series_from_filename == 0: worksheet.write(1, col_offset + 2, 'P.D. [µm]')
                chart_pd = writer.book.add_chart({'type': 'scatter'})
                chart_pd.add_series({
                    'name': [sheet_name, 1, col_offset + 2],  # Nazwa serii z komórki 'P.D. [µm]'
                    'categories': [sheet_name, 3, col_offset, 3 + len(df) - 1, col_offset],  # Zakres dla osi X
                    'values': [sheet_name, 3, col_offset + 2, 3 + len(df) - 1, col_offset + 2],  # Zakres dla osi Y
                    'line': {'width': 2},  # Ciągła linia bez punktów
                    'marker': {'type': 'none'},  # Wyłącza wyświetlanie punktów
                })
                chart_pd.set_title({'name': f"{cleaned_filename} - Penetration Depth [µm]"})
                chart_pd.set_x_axis({
                    'name': x_axis_label,  # nazwa osi X zależna od języka
                    'min': 0,  # minimalna wartość osi X
                    'max': 100 if x_axis_max < 10 else x_axis_max,  # zakres
                    'major_unit': max(25, int(round(x_axis_max / 5.0))) if x_axis_max >= 10 else 2,  # podziałka
                    'major_gridlines': {'visible': True},  # dodaj pionowe kreski
                })
                chart_pd.set_y_axis({
                    'name': y_axis_label_pd,  # nazwa osi Y zależna od języka
                })
                worksheet.insert_chart(f"J{25 + row_offset}", chart_pd)

            # Dodaj pustą kolumnę jako separator
            col_offset += len(df.columns) + 1
            # Aktualizuj offset wiersza
            row_offset += 15
    print(f"Dane zostały zapisane do pliku {output_xlsx}")

def generate_combined_xlsx_2(csv_files=None, csv_files_raw=None, output_xlsx="default.xlsx", series_from_filename=0, chart_lang='en'):
    # Sprawdź, czy przynajmniej jedno z csv_files lub csv_files_raw jest przekazane
    if not csv_files and not csv_files_raw:
        raise ValueError("At least one of 'csv_files' or 'csv_files_raw' must be provided.")
    
    with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
        col_offset = 0  # Przesunięcie kolumn na dane do wykresu (współrzędne do wykresów)
        row_offset = 0  # Przesunięcie wiersza na wykresy
        sheet_name = output_xlsx.replace('.xlsx', "")[:31]  # Nazwa arkusza danych (limit do 31 znaków)

        # Ustawienie nazw osi w zależności od języka
        if chart_lang == 'pl':
            x_axis_label = 'Droga [m]'
            y_axis_label_mu = 'Współczynnik tarcia'
            y_axis_label_pd = 'Zużycie liniowe [µm]'
        else:  # Domyślnie angielski
            x_axis_label = 'Distance [m]'
            y_axis_label_mu = 'Friction coefficient'
            y_axis_label_pd = 'Linear wear [µm]'

        for csv_file, csv_file_raw in zip(csv_files, csv_files_raw):
            # Wczytaj dane z głównego pliku CSV
            df = pd.read_csv(csv_file)
            # Wczytaj dane z pliku surowego CSV
            df_raw = pd.read_csv(csv_file_raw)

            # Zapisz dane z obu DataFrame do arkusza, z uwzględnieniem przesunięcia kolumn
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=col_offset)
            df_raw.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=col_offset + len(df.columns) + 1)

            # Uzyskaj nazwę pliku i tekst w nawiasach
            filename = os.path.basename(csv_file)
            match = re.search(r'\((.*?)\)', filename)
            text_in_brackets = match.group(1) if match else ""
            cleaned_filename = filename.replace('.csv', "").replace('(', "").replace(')', "")

            # Obiekt worksheet
            worksheet = writer.sheets[sheet_name]
            # Informacje na początku pliku .xlsx
            worksheet.write(0, col_offset, cleaned_filename)

            series_name = text_in_brackets if match else filename
            if series_from_filename == 1:
                if csv_files_raw: # jeżeli jest csv_files_raw
                    worksheet.write(1, col_offset + len(df.columns) + 2, series_name + " RAW")
                    worksheet.write(1, col_offset + len(df.columns) + 3, series_name + " RAW")
                if csv_files: # jeżeli jest csv_files
                    worksheet.write(1, col_offset + 1, series_name)
                    worksheet.write(1, col_offset + 2, series_name)

            # Maksymalna wartość dla osi X
            x_axis_max = int(round(df['Distance [m]'].max()))
            x_axis_max = (int((x_axis_max + 24) / 25)) * 25

            # Dodaj wykres µ
            chart = writer.book.add_chart({'type': 'scatter'})
            # Dodaj serię z danych surowych na tym samym wykresie RAW
            if series_from_filename == 0: worksheet.write(1, col_offset + len(df.columns) + 2, 'µ RAW')
            chart.add_series({
                'name': [sheet_name, 1, col_offset + len(df.columns) + 2],
                'categories': [sheet_name, 3, col_offset + len(df.columns) + 1, 3 + len(df_raw) - 1, col_offset + len(df.columns) + 1],
                'values': [sheet_name, 3, col_offset + len(df.columns) + 2, 3 + len(df_raw) - 1, col_offset + len(df.columns) + 2],
                'line': {'width': 2},
                'marker': {'type': 'none'},
            })
            # Dodaj serię z danych przefiltrowanych na tym samym wykresie MAIN
            if series_from_filename == 0: worksheet.write(1, col_offset + 1, 'µ')
            chart.add_series({
                'name': [sheet_name, 1, col_offset + 1],
                'categories': [sheet_name, 3, col_offset, 3 + len(df) - 1, col_offset],
                'values': [sheet_name, 3, col_offset + 1, 3 + len(df) - 1, col_offset + 1],
                'line': {'width': 2},
                'marker': {'type': 'none'},
            })
            chart.set_title({'name': f"{cleaned_filename} - µ"})
            chart.set_x_axis({
                'name': x_axis_label,
                'min': 0,
                'max': 100 if x_axis_max < 10 else x_axis_max,
                'major_unit': max(25, int(round(x_axis_max / 5.0))) if x_axis_max >= 10 else 2,
                'major_gridlines': {'visible': True},
            })
            chart.set_y_axis({
                'name': y_axis_label_mu,
            })
            worksheet.insert_chart(f"B{25 + row_offset}", chart)

            # Dodaj wykres Penetration Depth [µm] (jeśli istnieje)
            if 'Penetration Depth [µm]' in df.columns:
                chart_pd = writer.book.add_chart({'type': 'scatter'})
                # Dodaj serię z danych surowych na tym samym wykresie RAW
                if series_from_filename == 0: worksheet.write(1, col_offset + len(df.columns) + 3, 'P.D. RAW [µm]')
                chart_pd.add_series({
                    'name': [sheet_name, 1, col_offset + len(df.columns) + 3],
                    'categories': [sheet_name, 3, col_offset + len(df.columns) + 1, 3 + len(df_raw) - 1, col_offset + len(df.columns) + 1],
                    'values': [sheet_name, 3, col_offset + len(df.columns) + 3, 3 + len(df_raw) - 1, col_offset + len(df.columns) + 3],
                    'line': {'width': 2},
                    'marker': {'type': 'none'},
                })
                # Dodaj serię z danych przefiltrowanych na tym samym wykresie MAIN
                if series_from_filename == 0: worksheet.write(1, col_offset + 2, 'P.D. [µm]')
                chart_pd.add_series({
                    'name': [sheet_name, 1, col_offset + 2],
                    'categories': [sheet_name, 3, col_offset, 3 + len(df) - 1, col_offset],
                    'values': [sheet_name, 3, col_offset + 2, 3 + len(df) - 1, col_offset + 2],
                    'line': {'width': 2},
                    'marker': {'type': 'none'},
                })
                chart_pd.set_title({'name': f"{cleaned_filename} - Penetration Depth [µm]"})
                chart_pd.set_x_axis({
                    'name': x_axis_label,
                    'min': 0,
                    'max': 100 if x_axis_max < 10 else x_axis_max,
                    'major_unit': max(25, int(round(x_axis_max / 5.0))) if x_axis_max >= 10 else 2,
                    'major_gridlines': {'visible': True},
                })
                chart_pd.set_y_axis({
                    'name': y_axis_label_pd,
                })
                worksheet.insert_chart(f"J{25 + row_offset}", chart_pd)

            col_offset += len(df.columns) + len(df_raw.columns) + 2
            row_offset += 15
    print(f"Dane zostały zapisane do pliku {output_xlsx}")

def main():
    # Disable generation .pyc files
    sys.dont_write_bytecode = True
    # Enable ANSI escape codes in terminal
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    print("Program do obróbki danych z tribometru: TRB3, Nano TRB, T11, Rtec")
    print("Wczytuje pliki .txt i .csv z folderu z programem")
    print("Pliki te powinny mieć w nazwie nawias () a w nim nazwę serii")
    print("Pliki z Rtec powinny mieć w nazwie prędkość liniową np. 0.1m-s a dla T11 dodatkowo obciążenie np. 10N\n")
    folder_path = '.'
    csv_files = []
    csv_files_raw = []
    config_path = "config.ini"
    if not os.path.isfile(config_path):
        config_path = "_config.ini"
    if os.path.isfile(config_path):
        config = load_config(config_path) # Wczytaj dane z configu
        offset_raw = config['offset_raw']
        title_from_text = config['title_from_text']
        erase_peak = config['erase_peak']
        invert_peak = config['invert_peak']
        default_window_length_u = config['default_window_length_u']
        default_window_length_pd = config['default_window_length_pd']
        min_sample = config['min_sample']
        max_sample = config['max_sample']
        chart_lang = config['chart_lang']
    else:
        min_sample, max_sample, default_window_length_u, default_window_length_pd, title_from_text, offset_raw, erase_peak, invert_peak, chart_lang = ask_user_for_variables() # Wczytaj dane od użytkownika
    for filename in os.listdir(folder_path):
        if filename.endswith(".txt") or filename.endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            data, tribometer_type, mode = read_and_process_file(file_path) # Wczytaj dane z plików
            if data is not None:
                # Przetwórz dane przy użyciu najlepszej wartości sample_average (do wyboru µ lub pd)
                best_sample_average_µ, best_sample_average_pd = find_optimal_samples_average(data, column_names=['Distance [m]', 'µ', 'Penetration Depth [µm]'], min_sample=min_sample, max_sample=max_sample)
                averaged_data = adjust_and_average_data(data, best_sample_average_µ, default_window_length_u, default_window_length_pd)
                
                # Korekta zużycia liniowego
                averaged_data = process_penetration_depth(averaged_data, data, 0.05, offset_raw, erase_peak, invert_peak) # df=averaged_data, data=data, percent=0.1, offset_raw=0, erase_peak=0, invert_peak=1

                # Approximate the last value
                approximated_data = approximate_last_measurement(averaged_data, data)

                # Nazwa plików CSV odpowiadających nazwie pliku tekstowego oryginalnego + spacja
                output_file = os.path.splitext(filename)[0] + ' .csv'
                output_file_raw = "raw_" + output_file # to samo co wyżej, ale z przedrostkiem "raw_"
                
                if "Nano" in tribometer_type: # "Nano" - usuń kolumnę pd dla nano tribometru
                    approximated_data = approximated_data.drop(columns=['Penetration Depth [µm]']) # z obrobionych
                    data = data.drop(columns=['Penetration Depth [µm]']) # z raw

                # Zapisz wynik do pliku CSV
                try:
                    approximated_data.to_csv(output_file, index=False, float_format='%.4f') # finalne dane wyjściowe
                    data.to_csv(output_file_raw, index=False, float_format='%.4f') # dane tylko wstępnie obrobione
                    csv_files.append(output_file)
                    csv_files_raw.append(output_file_raw)
                    mean_u = display_average(approximated_data, 'µ')
                    print(f"[{tribometer_type}][{mode}] średnia dla µ: {mean_u:.4f}, plik: {filename}\n")
                except Exception as e:
                    print(f"\033[91mBłąd podczas zapisywania pliku {output_file}: {e} \033[0m")
                
    # Generowanie pliku xlsx ze wszystkimi danymi
    status = 0
    if csv_files or csv_files_raw: print("Generowanie pliku Excelowskiego ...")
    if csv_files:
        try:
            generate_combined_xlsx(csv_files, "combined_data.xlsx", title_from_text, chart_lang)
            status = 1
        except Exception as e:
            status = 0
            print(f"\033[91mBłąd podczas zapisywania pliku combined_data.xlsx: {e} \033[0m")
        if csv_files and csv_files_raw:
            try:
                generate_combined_xlsx_2(csv_files, csv_files_raw, "combined_data_all.xlsx", title_from_text, chart_lang)
                status = 1
            except Exception as e:
                status = 0
                print(f"\033[91mBłąd podczas zapisywania pliku combined_data_all.xlsx: {e} \033[0m")
        for file in csv_files:
            if os.path.exists(file):
                os.remove(file)
        # Sprawdź, czy wszystkie pliki zostały usunięte
        all_removed = True
        for file in csv_files:
            if os.path.exists(file):
                all_removed = False
                print(f"\033[38;5;214m Plik {file} nadal istnieje. \033[0m")
        if all_removed:
            print("Wszystkie pliki zostały usunięte.")
        # Generowanie pliku xlsx ze wszystkimi danymi
    if csv_files_raw:
        try:
            generate_combined_xlsx(csv_files_raw, "combined_data_raw.xlsx", title_from_text, chart_lang)
            status = 1
        except Exception as e:
            status = 0
            print(f"\033[91mBłąd podczas zapisywania pliku combined_data_raw.xlsx: {e} \033[0m")
        for file in csv_files_raw:
            if os.path.exists(file):
                os.remove(file)
        # Sprawdź, czy wszystkie pliki zostały usunięte
        all_removed = True
        for file in csv_files_raw:
            if os.path.exists(file):
                all_removed = False
                print(f"\033[38;5;214mPlik {file} nadal istnieje. \033[0m")
        if all_removed:
            print("Wszystkie pliki raw zostały usunięte.")
    if status == 1: print("\033[92mWszystkie dane z WCZYTANYCH plików zostały zapisane poprawnie \033[0m")
    else: print("\033[91mNie wszystkie dane zostały zapisane poprawnie z powodu powyższych błędów \033[0m")

    # Odliczanie do zamknięcia konsoli
    countdown = 5
    while countdown > 0:
        print(f"Okno konsoli zostanie zamknięte za {countdown} s, naciśnij ESC by anulować zamknięcie", end="\r")  # Wyświetl odliczanie w tej samej linii
        if msvcrt.kbhit() and msvcrt.getch() == b'\x1b':  # Sprawdź, czy naciśnięto klawisz Escape
            print("\n Anulowano zamknięcie, teraz naciśnij dowolny klawisz, aby zamknąć konsolę.")
            msvcrt.getch()  # Czeka na naciśnięcie dowolnego klawisza
        time.sleep(1)
        countdown -= 1

if __name__ == "__main__":
    main()

# TODO dodać komunikat że nie ma plików do wczytania (pusty folder)
# TODO poprawić komunikat braku zmiennych w nazwie pliku (usunąć nawias kwadratowy i standardowe zmienne 0.1m/s i 10N)
# TODO dodać procent przetwarzania danych
# TODO zunifikować funkcję generate_combined_xlsx i generate_combined_xlsx_2
