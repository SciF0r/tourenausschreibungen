# Requirements
python >=3.8

# Initialization
pytho -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

# Generate for Rote Karte
python collect.py rotekarte Tourenliste.xlsx

# Generate for Jahresprogramm
python collect.py jahresprogramm Tourenliste.xlsx
