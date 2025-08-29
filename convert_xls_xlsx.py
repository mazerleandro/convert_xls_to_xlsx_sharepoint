from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import pandas as pd
import os



# ================= CONFIGURAÇÕES =================
# Credenciais do Azure AD App Registration
client_id = "XXXXX"   # Application (client) ID
tenant_id = "XXXXXXXXX"   # Directory (tenant) ID
client_secret = "XXXXXXXXXXXXX"  # Client Secret

# URL do site do SharePoint
site_url = "https://XXXXXXXX.sharepoint.com/sites/VFIT_PowerBI"

# Caminho relativo do arquivo no SharePoint (sem %20 → use espaço normal)
file_relative_url = "/sites/VFIT_PowerBI/Shared Documents/POWERBI/INCs/INC_Report_PowerBI.xls"

# Nome local do arquivo baixado
local_xls = "INC_Report_PowerBI.xls"
local_xlsx = "INC_Report_PowerBI.xlsx"
# ==================================================


def baixar_arquivo():
    """Baixa o arquivo do SharePoint"""
    print("🔐 Autenticando no SharePoint...")
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

    print("⬇️ Baixando arquivo do SharePoint...")
    file = ctx.web.get_file_by_server_relative_url(file_relative_url)
    with open(local_xls, "wb") as f:
        file.download(f).execute_query()

    print(f"✅ Arquivo baixado: {local_xls}")


def converter_xls_para_xlsx():
    """Converte XLS para XLSX"""
    print("🔄 Convertendo XLS para XLSX...")
    df = pd.read_excel(local_xls, sheet_name=None)  # Lê todas as abas

    # Salva como XLSX
    with pd.ExcelWriter(local_xlsx, engine="openpyxl") as writer:
        for sheet, data in df.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

    print(f"✅ Conversão concluída: {local_xlsx}")


if __name__ == "__main__":
    baixar_arquivo()

    converter_xls_para_xlsx()
