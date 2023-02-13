import sqlite3
conn = sqlite3.connect('data.db',check_same_thread=False)
c = conn.cursor()

def create_table_prestation():
    c.execute("CREATE TABLE IF NOT EXISTS prestationtable(N° de prix  TEXT,Désignation TEXT,Unité TEXT,Type prestation TEXT,Temps Main d œuvre en heures REAL,Prix1 unitaire MO JOUR Taux horaire REAL,Prix2 unitaire MO JOUR Taux forfaitaire REAL,Prix3 unitaire MO NUIT COURTE Taux horaire REAL,Prix4 unitaire MO NUIT COURTE Taux forfaitaire REAL,Prix5 unitaire MO NUIT LONGUE Taux horaire REAL,Prix6 unitaire MO NUIT LONGUE Taux forfaitaire REAL)")

def add_data(a,b,c,d,f,g,h,i,j,k,l):
    c.execute("INSERT INTO prestationtable(N° de prix ,Désignation ,Unité ,Type prestation ,Temps Main d œuvre en heures ,Prix unitaire MO JOUR (hors fourniture)Taux horaire ,Prix unitaire MO JOUR (hors fourniture)Taux forfaitaire ,Prix unitaire MO NUIT COURTE (hors fourniture)Taux horaire ,Prix unitaire MO NUIT COURTE (hors fourniture)Taux forfaitaire ,Prix unitaire MO NUIT LONGUE (hors fourniture)Taux horaire ,Prix unitaire MO NUIT LONGUE (hors fourniture)Taux forfaitaire ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",(a,b,c,d,f,g,h,i,j,k,l))
    conn.commit()    
    
def create_table_pres():
    c.execute('CREATE TABLE IF NOT EXISTS prestationtable(task TEXT,task_status TEXT,task_due_date DATE)')