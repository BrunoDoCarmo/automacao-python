import pywhatkit as kit
from time import sleep
import pyautogui
import os
import pandas as pd
from datetime import datetime, timedelta
import locale

# Configurar localidade para formato de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def enviar_mensagem(nome, telefone, mensagem):
    try:
        print(f"Enviando a mensagem para {nome} ({telefone})...")

        # Enviar a mensagem instantaneamente
        kit.sendwhatmsg_instantly(telefone, mensagem)
        
        # Aguardar 5 segundos para garantir que a mensagem foi enviada
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')

        print(f'Mensagem enviada para {nome} ({telefone}) com sucesso!')

    except Exception as e:
        print(f"Ocorreu um erro ao enviar a mensagem para {nome} ({telefone}): {e}")

def formatar_telefone(telefone):
    telefone = str(telefone)
    # Remover espaços, parênteses e hífens
    telefone = telefone.replace(" ", "").replace("(", "").replace(")", "").replace("-", "")
    
    # Adicionar código do país se não estiver presente
    if not telefone.startswith("+55"):
        if telefone.startswith("+"):
            telefone = "+55" + telefone[1:]
        elif telefone.startswith("55"):
            telefone = "+" + telefone
        else:
            telefone = "+55" + telefone

    return telefone

def converter_valor(valor):
    try:
        valor = float(valor)
        valor_formatado = locale.currency(valor, grouping=True)
        return valor_formatado
    except ValueError:
        return valor

def converter_data(data):
    try:
        data_formatada = data.strftime('%d/%m/%Y')
        return data_formatada
    except Exception as e:
        print(f"Erro ao converter data: {e}")
        return data

# Ler a planilha de contatos
def main():
    planilha = 'contatos.xlsx'
    
    if not os.path.exists(planilha):
        print(f"Arquivo '{planilha}' não encontrado.")
        return
    
    df = pd.read_excel(planilha)

    # Debug: imprimir as colunas encontradas na planilha
    print("Colunas encontradas na planilha:", df.columns)

    # Validar se as colunas necessárias estão presentes
    required_columns = {'Nome', 'Telefone', 'Boleto', 'Vencimento', 'NFSe', 'Classificacao'}
    if not required_columns.issubset(df.columns):
        print(f"A planilha deve conter as colunas: {required_columns}")
        return
    
    # Iterar sobre cada linha da planilha e enviar a mensagem
    for index, row in df.iterrows():
        nome = row['Nome']
        telefone = formatar_telefone(row['Telefone'])
        boleto = converter_valor(row['Boleto'])
        vencimento = converter_data(pd.to_datetime(row['Vencimento']))
        nfse = row['NFSe']
        classificacao = row['Classificacao']
        # Formatando a mensagem
        if classificacao.lower() == 'mensalidade':
            if nfse.lower() == 'sim':
                mensagem = (
                    f'Prezado(a) {nome},\n\n'
                    f'Bom dia!\n\n'
                    f'Informamos que o boleto no valor de {boleto} com vencimento para o dia {vencimento} e a Nota Fiscal de Serviços Eletrônica (NFSe) referente à {classificacao} do mês serão enviados em breve.\n\n'
                    f'Agradecemos pela atenção e estamos à disposição para qualquer dúvida.\n\n'
                    f'Atenciosamente,\n\n'
                    f'Vivi, inteligência artificial da BDS Sistema'
                )
            else: 
                mensagem = (
                    f'Prezado(a) {nome},\n\n'
                    f'Bom dia!\n\n'
                    f'Informamos que o boleto no valor de {boleto} com vencimento para o dia {vencimento} referente à {classificacao} do mês serão enviados em breve.\n\n'
                    f'Agradecemos pela atenção e estamos à disposição para qualquer dúvida.\n\n'
                    f'Atenciosamente,\n\n'
                    f'Vivi, inteligência artificial da BDS Sistema'
                )
        elif classificacao.lower() == 'anuidade':
            if nfse.lower() == 'sim':
                mensagem = (
                    f'Prezado(a) {nome},\n\n'
                    f'Bom dia!\n\n'
                    f'Informamos que o boleto no valor de {boleto} com vencimento para o dia {vencimento} e a Nota Fiscal de Serviços Eletrônica (NFSe) referente à {classificacao} do mês serão enviados em breve.\n\n'
                    f'Agradecemos pela atenção e estamos à disposição para qualquer dúvida.\n\n'
                    f'Atenciosamente,\n\n'
                    f'Vivi, inteligência artificial da BDS Sistema'
                )
            else: 
                mensagem = (
                    f'Prezado(a) {nome},\n\n'
                    f'Bom dia!\n\n'
                    f'Informamos que o boleto no valor de {boleto} com vencimento para o dia {vencimento} referente à {classificacao} do mês serão enviados em breve.\n\n'
                    f'Agradecemos pela atenção e estamos à disposição para qualquer dúvida.\n\n'
                    f'Atenciosamente,\n\n'
                    f'Vivi, inteligência artificial da BDS Sistema'
                )
        else:
            print(f"Classificação desconhecida na linha {index + 1}, pulando este contato.")
            continue
        
        # Enviar mensagem
        enviar_mensagem(nome, telefone, mensagem)

        # Limpar o terminal (no Windows)
        os.system('cls' if os.name == 'nt' else 'clear')

if __name__ == "__main__":
    main()
