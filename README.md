# Verificare subvenție — calcul independent

Aplicație [Streamlit](https://streamlit.io/) care verifică cheltuielile dintr-o subvenție,
**recalculând totul singură, pas cu pas**, fără să se bazeze pe formulele din Excel (care pot fi greșite).
Poate analiza raportul de subvenție, extrasul de cont bancar, sau ambele — caz în care le și compară între ele.

## Ce face

### Raportul de subvenție
- Detectează automat structura fișierului (coloane fixe + categorii deduse automat).
- Recalculează totalurile pornind doar de la rândurile de tranzacții reale.
- **Harta celulelor cu probleme**: recalculează fiecare celulă derivată (Total cheltuită, Sold cont, rândul de Total)
  și marchează unde lipsesc formule sau unde sunt valori hardcodate care nu respectă formula.
- Raport de **calitate a datelor**: celule lipsă, ani greșiți, incoerențe dată ↔ factură.
- Export Excel curat + fișier cu celulele problematice marcate pe hartă (roșu = hardcodat greșit, galben = formulă lipsă).

### Extrasul de cont bancar (sursa reală de adevăr)
- Calculează totul din debit/credit: sold inițial → + intrări → − ieșiri → sold final, cu auto-verificare
  față de ce declară banca (sold final, rulaje).
- Defalcare pe luni, pe zile, pe tip de operațiune și pe beneficiar.

### Reconciliere raport ↔ extras
- Potrivire plată-cu-plată (sumă + furnizor): ce plată din bancă lipsește sau diferă față de raport.
- Verificarea soldului „Sold cont” (AG) față de soldul real din bancă.
- **Raport de diferențe pe zi** (intrări / ieșiri / sold), cu comisioanele luate în calcul de ambele părți.

## Rulare locală

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt
streamlit run abc.py
```

Apoi încarci raportul de subvenție și/sau extrasul de cont (.xls / .xlsx).

## Fișiere
- `abc.py` — aplicația completă (raport + extras + reconciliere).
- `varianta_veche.py` — varianta inițială, doar analiza raportului de subvenție.
- `requirements.txt` — dependențele.

## Publicare online (live)
Aplicația poate fi pusă live gratuit pe [Streamlit Community Cloud](https://share.streamlit.io/):
conectezi contul GitHub, alegi acest repo și fișierul `abc.py`.
