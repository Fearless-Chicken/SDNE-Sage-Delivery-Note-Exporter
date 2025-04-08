@echo off
echo [*] Installation des dépendances Python...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo [✅] Installation terminée.
pause
