# ────────────────────────────────────
# SDNE - Sage Delivery Note Exporter
# Développé par Fearless-Chicken
# © 2025 - Tous droits réservés
# ────────────────────────────────────
import pyodbc, csv
from typing import List, Tuple, Any, Optional
from datetime import date

## Définition de variables utiles
headers = ["Type de document","Ref client","Date de départ","Nom destinataire","Adresse","Code postal","Ville","Mobile","Mail"]
NumberToMonth = {1: "Jan",2: "Fev",3: "Mar",4: "Avr",5: "Mai",6: "Jun",7: "Jul",8: "Aou",9: "Sep",10: "Oct",11: "Nov",12: "Dec"}

def correcTel(num:str)->str:
    '''Sert à ré-arranger le format des numéro de tel'''
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
    '''Remet des espaces entre les numéro de tel'''
    return ' '.join(f"{num[i]}{num[i+1]}" for i in range(0, len(num), 2))

def extractConf(config = {}) -> dict[str:str]:
    '''Récupère la config dans le fichier de conf '''
    with open('app/config/config.conf', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        config['server']   = lines[0].strip().split(" : ")[1].replace('"', '').replace("'", '')
        config['database'] = lines[1].strip().split(" : ")[1].replace('"', '').replace("'", '')
        config['username'] = lines[2].strip().split(" : ")[1].replace('"', '').replace("'", '')
        config['password'] = lines[3].strip().split(" : ")[1].replace('"', '').replace("'", '')
    return config

def con(conf=extractConf()) -> Tuple:
    '''Se sert de la configuration pour se connecter à la database'''
    print("🔄 Connection à la base de donnée...")
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
        print("❌ Erreur lors de la connection")
        return None, None
    print("✅ Connection établie..")
    return conn, cursor

def ExtractD2D(cursor, dateDeb, dateFin):
    '''Exécute la requête formatée avec les dates fournies'''
    cursor.execute(open('app/SQL/ExtractD2D.sql','r',encoding='utf-8').read().replace("{dateDeb}", dateDeb).replace("{dateFin}", dateFin))
    return cursor

def ExtractALL(cursor):
    '''Exécute la requête extractALL'''
    cursor.execute(open('app/SQL/ExtractAllUnexported.sql','r',encoding='utf-8').read())
    return cursor

def tryDate(date:str) -> Optional[str]:
    '''Test si le format de la date est correcte pour la rendre en YYYY-MM-DD'''
    from datetime import datetime    
    try:
        # Essaye de parser au format jj-mm-aaaa
        date_obj = datetime.strptime(date, "%d-%m-%Y")
        # Retourne au format yyyy-mm-dd
        return date_obj.strftime("%Y-%m-%d")
    except ValueError:
        input("❌ Format invalide. Utilise jj-mm-aaaa.\nAppuyer sur <entrer> pour continuer...")
        return None

def clear()->None:
    '''Vide l'écran'''
    import os
    os.system('cls' if os.name == 'nt' else 'clear')

def log(result:str, dateDeb:date, dateFin:date, fileName:str, nb:int, errorCSV:str, errorSQL:str) -> None:
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

def emptyLog()->None:
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
        f.write(f"""{curLogs}\n[{datetime.today().strftime('%d-%m-%Y_%H:%M:%S')}] Pas de nouveaux BL, Aucun export effectué
    -------------------------------\n""")

def setExtracted(ListeBL:list,cursor)-> Optional[str]:
    from datetime import datetime
    try:
        print("🔄 Modification de l'état des BL dans la base")
        for BL in ListeBL:
            request = open("app/SQL/UpdateBL.sql","r",encoding="utf-8").read().replace("{todaydate}", datetime.today().strftime('%Y-%m-%d %H:%M:%S')).replace("{BL}",BL)
            cursor.execute(request)
            cursor.commit()
        print("✅ Modification de l'état des BL dans la base terminé avec succès")
        return None
    except Exception as e:
        errorSQL = str(e)
        print("❌ Erreur lors de la modification !")
        return errorSQL
    
# def CreateCSV(data:list,nb=0,error=None,listeBL=[])->Tuple[str, date, date, str, int, str, List[Any]]:
def CreateCSV(data:list,nb=0,error=None,listeBL=[])->Tuple[str, str, int, str, List[Any]]:
    '''1. Créer un CSV <D2D_export_bl_JJ-MM-AAAA_hh-mm-ss.csv> à partir des données récupérées
       2. Ajoute le tag exporté dans la base de données pour éviter les doublons
       3. Déclanche la gestion des logs'''
    import os
    if "extracted" not in os.listdir("app"):
        os.mkdir("app/extracted")
    try:
        print("🔄 Exportation des BL.")
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
        print("❌ Erreur lors de l'export !")
        error = str(e)
    # return "OK" if error is None else "KO", dateDeb, dateFin, f"D2D_export_bl_{datetime.today().strftime('%d-%m-%Y_%H-%M-%S')}.csv", nb, error, listeBL
    return "OK" if error is None else "KO", f"D2D_export_bl_{datetime.today().strftime('%d-%m-%Y_%H-%M-%S')}.csv", nb, error, listeBL

if __name__ == "__main__":
    dateDeb = "Export"
    dateFin = "total"
    try: 
        ## Connection à la base de données
        conn, cursor = con()
        if conn is None:
            raise Exception("Connection failed")

        ## Exécution de la requête SQL
        try:
            cursor = ExtractALL(cursor)
        except Exception as e:
            print("❌ Erreur dans la requête :")        
    
        ## Récupération de la data exporté
        data = cursor.fetchall()
        if data != []:
            ## Création du CSV
            status, name, nb, errorCSV, listeBL = CreateCSV(data)
            print(f"✅ {nb} Bon{"s" if nb>1 else""} de livraison{"s" if nb>1 else""} {"ont" if nb>1 or nb==0 else "a"} été exporté{"s" if nb>1 else""} {"" if nb == 0 else "avec succès."}")

            ## Mise à jour du status des BL dans la base
            errorSQL = setExtracted(listeBL, cursor)
            
            ## Création des logs de la session
            log(status, dateDeb, dateFin, name, nb, errorCSV, errorSQL)
            
        else:
            emptyLog()
            print("📭 Aucun nouveau bon de livraison, pas d'export nécessaire.")
        conn.close()
        input("✅ Vous pouvez fermer la fenêtre ou appuyer sur <Entrée> pour terminer.")
    except KeyboardInterrupt:
        clear()
        input("Appuyer sur <entrer> pour quitter...")
        clear()
    except Exception as e:
        print(e)