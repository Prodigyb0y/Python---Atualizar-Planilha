import time
import win32com.client as win32

# Caminho para a planilha
caminho_planilha = r"caminho_planilha"


# Crie uma instância do Excel
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False  # Não exibir a interface do Excel

print("\nIniciando Atulização Planilha {caminho_planilha})...")
# Abra a planilha
print("Aguardando abrir planilha...")
workbook = excel.Workbooks.Open(caminho_planilha)
time.sleep(5)

# Atualize todas as consultas
print("Atualizando Querys...")
workbook.RefreshAll()
excel.CalculateUntilAsyncQueriesDone() #Aguarda atualizações finalizarem

# Salve a planilha
workbook.Save()
time.sleep(5)
workbook.Close()
print("Planilha Salva!")

print("Processo: {caminho_planilha} finalizado")

time.sleep(5)


