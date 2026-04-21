import fitz  # pip install pymupdf
import os
import re

# --- CONFIGURAÇÕES ---
CAMINHO_ASSINATURA = 'P:\\Geral-pCloud\\02_FRANÇA\\01_ERT\\03 Assinatura +INNOVATION\\Assinatura.png' # A tua imagem (fundo transparente é melhor)
BASE_PATH = r"P:\\Geral-pCloud\\02_FRANÇA\\01_ERT\\2026\\02 ATTs FACT 2026"

def obter_caminho_trabalho():
    print("--- Configuração de Caminho ---")
    data_completa = input("Introduza o dia de envio (ex: 17-03-2026): ").strip()

    # Validar se o formato tem os traços necessários para extrair o mês
    # Esperamos algo como DD-MM-YYYY, logo o mês está entre o primeiro e o segundo traço
    partes = data_completa.split("-")

    if len(partes) != 3:
        print("\n❌ ERRO: Formato de data inválido. Use o formato DD-MM-YYYY.")
        return None
    
    # Extrair o mês e o ano para a pasta intermédia (ex: "03-2026")
    mes_ano = f"{partes[1]}-{partes[2]}"
    
    caminho_final = os.path.join(BASE_PATH, mes_ano, data_completa)
    
    if not os.path.exists(caminho_final):
        print(f"\n❌ ERRO: O caminho não foi encontrado:\n{caminho_final}")
        return None
    
    print(f"\n📂 Pasta selecionada: {caminho_final}\n")
    return caminho_final

def modulo_1_assinatura_att(pasta):
    print("--- MÓDULO 1: Assinatura de ATTs ---")
    for nome in os.listdir(pasta):
        if nome.upper().startswith("ATT") and nome.lower().endswith(".pdf"):
            caminho = os.path.join(pasta, nome)
            try:
                doc = fitz.open(caminho)
                for pagina in doc:
                    instancias = pagina.search_for("Le prestataire")
                    for inst in instancias:
                        area = fitz.Rect(inst.x0 - 100, inst.y1 + 15, inst.x0 + 150, inst.y1 + 75)
                        pagina.insert_image(area, filename=CAMINHO_ASSINATURA)
                doc.save(caminho, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
                doc.close()
                print(f"✅ ATT Assinada: {nome}")
            except Exception as e:
                print(f"❌ Erro na ATT {nome}: {e}")

def modulo_2_completar_fatura(pasta):
    print("\n--- MÓDULO 2: Anexar ATTs às Faturas (Substituição) ---")
    todos_ficheiros = os.listdir(pasta)
    
    for nome_fa in todos_ficheiros:
        if nome_fa.upper().startswith("FACTURE") and nome_fa.lower().endswith(".pdf"):
            caminho_fa = os.path.join(pasta, nome_fa)
            
            try:
                doc_fa = fitz.open(caminho_fa)
                texto_fatura = ""
                for pagina in doc_fa:
                    texto_fatura += pagina.get_text()

                match = re.search(r"ATT[\w_-]+", texto_fatura)
                
                if match:
                    nome_att_referenciada = match.group(0)
                    ficheiro_att = next((f for f in todos_ficheiros if nome_att_referenciada in f and f.lower().endswith(".pdf")), None)
                    
                    if ficheiro_att:
                        caminho_att = os.path.join(pasta, ficheiro_att)
                        doc_att = fitz.open(caminho_att)
                        
                        # Inserir a ATT na página 2
                        doc_fa.insert_pdf(doc_att, from_page=0, to_page=0, start_at=1)
                        
                        # --- LÓGICA DE SUBSTITUIÇÃO ---
                        caminho_temp = caminho_fa + ".tmp"
                        doc_fa.save(caminho_temp)
                        doc_fa.close()
                        doc_att.close()
                        
                        # Remove a original e renomeia a temporária para o nome original
                        os.remove(caminho_fa)
                        os.rename(caminho_temp, caminho_fa)
                        
                        print(f"✅ Fatura atualizada e substituída: {nome_fa}")
                    else:
                        doc_fa.close()
                        print(f"⚠️ ATT '{nome_att_referenciada}' não encontrada para {nome_fa}")
                else:
                    doc_fa.close()
                    print(f"🔍 Nenhuma referência ATT em {nome_fa}")
                
            except Exception as e:
                print(f"❌ Erro na Fatura {nome_fa}: {e}")

if __name__ == "__main__":
    # Primeiro assinamos tudo, depois anexamos
    pasta_trabalho = obter_caminho_trabalho()
    
    if pasta_trabalho:
        modulo_1_assinatura_att(pasta_trabalho)
        modulo_2_completar_fatura(pasta_trabalho)
        print("\n✨ Processo concluído!")