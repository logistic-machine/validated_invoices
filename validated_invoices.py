import os
import time
import xlwings as xw
import win32com.client as win32
from datetime import datetime

PASTA_BASE = 

DADOS_ENVIO = 
  
print(f"🚀 Iniciando automação em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

try:
    
    excel_app = xw.App(visible=True, add_book=False)
    outlook_app = win32.Dispatch('outlook.application')
except Exception as e:
    print(f"❌ Erro ao iniciar aplicativos: {e}")
    exit()


for item in DADOS_ENVIO:
    caminho_anexo = os.path.join(PASTA_BASE, item["arquivo"])
    destinatarios = "; ".join(item["to"])
    
    print(f"\n🔹 Processando: {item['arquivo']}")
    

    if not os.path.exists(caminho_anexo):
        print(f"⚠️ ALERTA: Arquivo não encontrado: {caminho_anexo}")
        continue

    try:
       
        print(f"   - Atualizando dados...")
        wb = excel_app.books.open(caminho_anexo)
        wb.api.RefreshAll()
        time.sleep(5) 
        wb.save()
        wb.close()
        print(f"   - Excel atualizado e salvo.")


        print(f"   - Preparando e-mail para {item['to'][0]}...")
        email = outlook_app.CreateItem(0)
        email.To = destinatarios
        email.Subject = "ERRATA: NOTAS VALIDADAS - CIF"
        
        email.HTMLBody = f"""
        <p>Prezados,</p>
        <p>Solicitamos a fineza de verificar se as notas já estão com vocês. Caso positivo, pedimos que seja realizada a entrada no estoque pelo Monitor Logístico, lembrando que as notas fiscais validadas constam como 'MONITOR LOGÍSTICO'.</p>
        <p>Reforçamos a importância do preenchimento correto da data real de recebimento dos materiais no Monitor Logístico. Esse cuidado é essencial porque a data registrada é a referência utilizada para calcular o prazo de pagamento aos fornecedores. Qualquer incorreção pode causar grandes problemas financeiros, como atrasos ou divergências no faturamento, prejudicando o relacionamento com nossos parceiros. Portanto, solicitamos máxima atenção e rigor nesse processo. Garantir informações precisas é fundamental para o bom andamento de nossas operações.</p>
        <p>Observação: Caso haja alguma nota fiscal já recebida que não esteja constando no arquivo, não há problema. As notas levam cerca de 24h para serem importadas após a sua emissão.</p>
        """
        
        email.Attachments.Add(caminho_anexo)
        email.Send()
        print(f"   ✅ Sucesso: E-mail enviado!")

    except Exception as e:
        print(f"   ❌ Erro no processo de {item['arquivo']}: {e}")

excel_app.quit()
print(f"\n🏁 Processo finalizado com sucesso em {datetime.now().strftime('%H:%M:%S')}")