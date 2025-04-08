import pick
import pyodbc
import csv

headers = [
    "Type de document",
    "Ref client",
    "Date de départ",
    "Nom destinataire",
    "Adresse",
    "Code postal",
    "Ville",
    "Mobile",
    "Mail"
]

def correcTel(num:str)->str:
    num = num.replace(" ", "").replace(".", "").replace("-", "")
    if len(num) == 10 and num[0] == '0':
        return num
    elif len(num) == 9 and num[0] != '0':
        return "0" + num
    elif len(num) == 12 and num[0:3] == '+33':
        return "0" + num[3:]
    elif len(num) == 13 and num[0:4] == '0033':
        return "0" + num[4:]
    else:
        return "Numéro de téléphone invalide"

def spaceTel(num:str)->str:
    return ' '.join(f"{num[i]}{num[i+1]}" for i in range(0, len(num), 2))

def extractConf(config = {}):
    with open('../config.conf', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        config['server']   = lines[0].strip().split(" : ")[1].replace('"', '').replace("'", '')
        config['database'] = lines[1].strip().split(" : ")[1].replace('"', '').replace("'", '')
        config['username'] = lines[2].strip().split(" : ")[1].replace('"', '').replace("'", '')
        config['password'] = lines[3].strip().split(" : ")[1].replace('"', '').replace("'", '')
    return config

def con(conf=extractConf()):
    conn_str = (
        'DRIVER={ODBC Driver 17 for SQL Server};'
        f'SERVER={conf['server']};'
        f'DATABASE={conf['database']};'
        f'UID={conf['username']};'
        f'PWD={conf['password']}'
    )
    try:
        conn = pyodbc.connect(conn_str)
        if conn is None:
            print("nop")
        cursor = conn.cursor()
    except Exception as e:
        print("err",e)
        return None, None
    return conn, cursor


def GenReq(dateDeb, dateFin):
    return f"""
SELECT 
    CASE [BIJOU].[dbo].[F_DOCENTETE].[DO_Type] 
        WHEN 3 THEN 'BL'
        ELSE 'AUTRE'
    END AS [Identifiant du segment],
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Tiers] AS [Ref client],
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Date] AS [Date de départ],
    [BIJOU].[dbo].[F_COMPTET].[CT_Intitule] AS [Nom destinataire],
    [BIJOU].[dbo].[F_COMPTET].[CT_Adresse] AS [Libellé de voie],
    [BIJOU].[dbo].[F_COMPTET].[CT_CodePostal] AS [Code postal],
    [BIJOU].[dbo].[F_COMPTET].[CT_Ville] AS [Localité],
    contact.[CT_TelPortable] AS [Mobile],
    contact.[CT_EMail] AS [Mail]
FROM 
    [BIJOU].[dbo].[F_DOCENTETE] 
JOIN 
    [BIJOU].[dbo].[F_COMPTET] 
    ON [BIJOU].[dbo].[F_DOCENTETE].[DO_Tiers] = [BIJOU].[dbo].[F_COMPTET].[CT_Num]
OUTER APPLY (
    SELECT TOP 1 *
    FROM [BIJOU].[dbo].[F_CONTACTT]
    WHERE [F_CONTACTT].[CT_Num] = [F_DOCENTETE].[DO_Tiers]
) AS contact
WHERE
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Type] = 3
	AND
	[BIJOU].[dbo].[F_DOCENTETE].[DO_Date] BETWEEN
		CAST('{dateDeb}T00:00:00.000' AS datetime)
		AND
		CAST('{dateFin}T23:59:59.997' AS datetime)
    ORDER BY 
        [BIJOU].[dbo].[F_DOCENTETE].[DO_Date] DESC
;
"""

def tryDate(date):
    from datetime import datetime    
    try:
        # Essaye de parser au format jj-mm-aaaa
        date_obj = datetime.strptime(date, "%d-%m-%Y")
        # Retourne au format yyyy-mm-dd
        return date_obj.strftime("%Y-%m-%d")
    except ValueError:
        input("❌ Format invalide. Utilise jj-mm-aaaa.\nAppuyer sur <entrer> pour continuer...")
        return None

def clear():
    import os
    os.system('cls' if os.name == 'nt' else 'clear')

def getInt(prompt, min, max):
    while True:
        try:
            clear()
            value = int(input(prompt))
            if min <= value <= max:
                return value
            else:
                print(f"❌ Valeur invalide. Doit être entre {min} et {max}.")
        except ValueError:
            print("❌ Format invalide. Utilise un entier.")

def CreateCSV(data,c):
    import os
    if "extracted" not in os.listdir():
        os.mkdir("extracted")

    from datetime import datetime
    with open(f"../extracted/{c}_export_bl_{datetime.today().strftime('%d-%m-%Y_%H-%M-%S')}.csv", "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(headers)
        for row in data:
            writer.writerow([
                row[0],                          # Type
                row[1],                          # Ref client
                row[2].strftime("%Y-%m-%d"),     # Date format1ée
                row[3], row[4], row[5], row[6],
                spaceTel(correcTel(row[7])) if row[7] else '',        # Mobile
                row[8] if row[8] else ''         # Mail
            ])
    print("✅ Export terminé avec succès.")

if __name__ == "__main__":
    try: 
        ## On tente de se connecter à la base de données
        print("Connecting to database...")
        conn, cursor = con()
        if conn is None:
            raise Exception("Connection failed")
        print("Connected to database successfully.")
        
        menu = ["1 - Date à date", "2 - Tout un mois","3 - Tout une année","4 - Quitter"]
        choix=pick.pick(title="Faites un choix", options=menu, indicator=">")[0]

        match choix:
            case "1 - Date à date":
                clear()
                dateDeb = None
                while dateDeb is None:
                    clear()
                    dateDeb = tryDate(input("Entre une date de début de période (jj-mm-aaaa) : "))
                clear()
                dateFin = None
                while dateFin is None:
                    clear()
                    dateFin = tryDate(input("Entre une date de fin de période (jj-mm-aaaa) : "))
                c = "D2D"
            case "2 - Tout un mois":
                mois = getInt("Entre le mois (1-12) : ", 1, 12)
                annee = getInt("Entre l'année (aaaa) : ", 2000, 2100)
                dateDeb = f"{annee}-{mois:02d}-01"
                fin = 31 if mois in [1, 3, 5, 7, 8, 10, 12] else 30 if mois != 2 else (29 if annee % 4 == 0 else 28)
                dateFin = f"{annee}-{mois:02d}-{fin}"
                c = "WM"
            case "3 - Tout une année":
                annee = getInt("Entre l'année (aaaa) : ", 2000, 2100)
                dateDeb = f"{annee}-01-01"
                dateFin = f"{annee}-12-31"
                c = "WY"

        ## On essaie d'exécuter la requête SQL
        try:
            cursor.execute(GenReq(dateDeb, dateFin))
        except Exception as e:
            print("Erreur dans la requête :", e)
        
        
        data = cursor.fetchall()
        # for row in rows:
        #     print(row)
        CreateCSV(data,c)
        conn.close()
        input("Vous pouvez fermer la fenêtre.")
    except KeyboardInterrupt:
        clear()
        input("Appuyer sur <entrer> pour quitter...")
        clear()
    except Exception as e:
        print(e)

