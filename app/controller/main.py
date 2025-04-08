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

NumberToMonth = {
        1: "Jan",
        2: "Fev",
        3: "Mar",
        4: "Avr",
        5: "Mai",
        6: "Jun",
        7: "Jul",
        8: "Aou",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dec"
}

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
    with open('app/config/config.conf', 'r', encoding='utf-8') as f:
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
    return open('app/SQL/extract.sql','r',encoding='utf-8').read().replace("{dateDeb}", dateDeb).replace("{dateFin}", dateFin)

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

def log(result, dateDeb, dateFin, fileName, nb, errorCSV, errorSQL):
    '''
    1. Déplace le dossier log<Ancien_Mois> s'il s'agit du premier export du mois 
    2. Ajoute les logs dans log<Mois_Actuel> avec le résultat de l'export
    '''
    from datetime import datetime
    import os

    ## Etape 1 
    listDir = os.listdir("app/log")
    listDir.remove("Ancienne logs")
    listDir.remove("template.log")
    if listDir[0] != f"log{NumberToMonth[int(datetime.today().strftime('%m'))]}.log":
        os.rename(
            os.path.join("log", listDir[0]),
            os.path.join("log", "Ancienne logs", listDir[0])
        )

    ## Etape 2
    try:
        curLogs = open(f'app/log/log{NumberToMonth[int(datetime.today().strftime('%m'))]}.log', 'r', encoding='utf-8').read()
    except FileNotFoundError:
        print("Le fichier de log de ce mois n'existe pas. Création en cours...")
        curLogs = ""
    with open(f'app/log/log{NumberToMonth[int(datetime.today().strftime('%m'))]}.log', 'w', encoding='utf-8') as f:
        f.write(f"""{curLogs}\n[{datetime.today().strftime('%d-%m-%Y_%H:%M:%S')}] Export effectué
    Résultat : {result}
    Plage : {dateDeb} -> {dateFin}
    Nom du fichier : {fileName}
    Nombre de BL : {nb}
    Liste des BL : {', '.join(listeBL)}
    ErreurCSV : {errorCSV}
    ErreurSQL : {errorSQL}
    -------------------------------\n""")

def setExtracted(ListeBL:list,cursor):
    from datetime import datetime
    try:
        for BL in ListeBL:
            request = f"""
UPDATE 
    [BIJOU].[dbo].[F_DOCENTETE]
SET 
	[BIJOU].[dbo].[F_DOCENTETE].[Exporté] = 'Oui',
	[BIJOU].[dbo].[F_DOCENTETE].[Date d'export] = '{datetime.today().strftime('%Y-%m-%d %H:%M:%S')}'
WHERE 
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Piece] = '{BL}';
"""
            cursor.execute(request)
            cursor.commit()
        return None
    except Exception as e:
        errorSQL = str(e)
        return errorSQL
    
def CreateCSV(data,nb=0,error=None,listeBL=[]):
    '''1. Créer un CSV <D2D_export_bl_JJ-MM-AAAA_hh-mm-ss.csv> à partir des données récupérées
       2. Ajoute le tag exporté dans la base de données pour éviter les doublons
       3. Déclanche la gestion des logs'''
    import os
    if "extracted" not in os.listdir("app"):
        os.mkdir("app/extracted")
    try:
        from datetime import datetime
        with open(f"app/extracted/D2D_export_bl_{datetime.today().strftime('%d-%m-%Y_%H-%M-%S')}.csv", "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(headers)
            for row in data:
                nb+=1
                writer.writerow([
                    row[0],                                               # Type
                    row[1],                                               # Ref BL
                    row[2],                                               # Ref client
                    row[3].strftime("%Y-%m-%d"),                          # Date formatée
                    row[4], row[5], row[6], row[7],
                    spaceTel(correcTel(row[8])) if row[8] else '',        # Mobile
                    row[9] if row[9] else ''                              # Mail
                ])
                listeBL.append(row[1])
        print("✅ Export terminé avec succès.")
    except Exception as e:
        print("❌ Erreur lors de l'écriture du fichier :", e)
        error = str(e)
    return "OK" if error is None else "KO", dateDeb, dateFin, f"D2D_export_bl_{datetime.today().strftime('%d-%m-%Y_%H-%M-%S')}.csv", nb, error, listeBL

if __name__ == "__main__":
    try: 
        ## On tente de se connecter à la base de données
        print("Connecting to database...")
        conn, cursor = con()
        if conn is None:
            raise Exception("Connection failed")
        print("Connected to database successfully.")
        
        ## On demande à l'utilisateur de rentrer les dates de début et de fin
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

        ## On essaie d'exécuter la requête SQL
        try:
            cursor.execute(GenReq(dateDeb, dateFin))
        except Exception as e:
            print("Erreur dans la requête :", e)
        
        
        data = cursor.fetchall()
    
        status, dateDeb, dateFin, name, nb, errorCSV, listeBL = CreateCSV(data)
        errorSQL = setExtracted(listeBL, cursor)
        log(status, dateDeb, dateFin, name, nb, errorCSV, errorSQL)
        conn.close()
        input("Vous pouvez fermer la fenêtre.")
    except KeyboardInterrupt:
        clear()
        input("Appuyer sur <entrer> pour quitter...")
        clear()
    # except Exception as e:
    #     print(e)