import pyodbc
import xlsxwriter

# Chemin vers le fichier Access
access_file_path = r'C:\chemin\vers\Database1 - Copie.accdb'
output_excel_path = r'C:\chemin\vers\output.xlsx'

# Connexion à la base de données Access
conn = pyodbc.connect(
    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + access_file_path + ';'
)
cursor = conn.cursor()

# Lister toutes les tables de la base
cursor.execute("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
tables = [row[0] for row in cursor.fetchall()]

# Créer un fichier Excel
workbook = xlsxwriter.Workbook(output_excel_path)

# Exporter chaque table vers une feuille Excel
for table in tables:
    worksheet = workbook.add_worksheet(name=table[:31])  # Limiter les noms des feuilles à 31 caractères
    cursor.execute(f"SELECT * FROM {table}")
    
    # Ajouter les colonnes (en-têtes)
    columns = [desc[0] for desc in cursor.description]
    for col_num, column_name in enumerate(columns):
        worksheet.write(0, col_num, column_name)
    
    # Ajouter les données
    for row_num, row in enumerate(cursor.fetchall(), start=1):
        for col_num, cell_value in enumerate(row):
            worksheet.write(row_num, col_num, cell_value)

# Fermer la connexion et le fichier Excel
conn.close()
workbook.close()

print(f"Conversion terminée. Fichier Excel : {output_excel_path}")
