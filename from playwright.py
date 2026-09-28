import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright

# Caminho de destino para salvar a planilha
PASTA_DESTINO = r"C:\Users\Micro\OneDrive\Área de Trabalho\Nova pasta (3)"

def raspar_produtos_em_promocao():
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=False)
        contexto = navegador.new_context()
        pagina = contexto.new_page()
        
        # Aumenta o tempo limite geral para evitar desconexões
        pagina.set_default_timeout(60000)
        
        # 1. Login no site
        print("Realizando login...")
        pagina.goto("https://www.megaleste.com.br/")
        pagina.fill("input[name='user']", "737275")
        pagina.fill("input[name='pass']", "1020")
        pagina.click("form[action*='/login/auth'] button[type='submit']")
        pagina.wait_for_load_state("networkidle")
        
        produtos_promocao = []
        seletor_card = ".product-line"
        total_paginas = 56  # Definido para varrer 57 páginas
        
        print(f"\n--- Iniciando Varredura de {total_paginas} Páginas ---")
        
        # 2. Iteração das 57 páginas
        for num_pagina in range(1, total_paginas + 1):
            url_pagina = f"https://www.megaleste.com.br/c/promocoes?page={num_pagina}"
            print(f"Acessando Página {num_pagina}/{total_paginas}...")
            
            try:
                pagina.goto(url_pagina)
                pagina.wait_for_selector(seletor_card, timeout=8000)
            except Exception:
                print(f" Página {num_pagina} demorou a responder ou não contém produtos. Pulando...")
                continue
                
            cards = pagina.locator(seletor_card).all()
            
            # 3. Extração dos dados da página atual
            for card in cards:
                try:
                    tem_promocao = card.locator(".price strike").count() > 0
                    if not tem_promocao:
                        continue
                    
                    nome = card.locator(".product-content h4").inner_text().strip()
                    codigo = card.locator(".product-content small").inner_text().strip()
                    
                    preco_antigo_texto = card.locator(".price strike").inner_text().strip()
                    preco_promo_texto = card.locator(".price span.text-danger").inner_text().strip()
                    
                    preco_limpo = (
                        preco_promo_texto.replace("R$", "")
                                         .replace(" ", "")
                                         .replace(".", "")
                                         .replace(",", ".")
                                         .strip()
                    )
                    valor_num = float(preco_limpo)
                    
                    produtos_promocao.append({
                        "Página": num_pagina,
                        "Código": codigo,
                        "Produto": nome,
                        "Preço Antigo": preco_antigo_texto,
                        "Preço Promocional": preco_promo_texto,
                        "Valor Numérico (R$)": valor_num
                    })
                    
                except Exception:
                    continue
            
            # Pequena pausa para estabilidade do servidor
            time.sleep(0.5)

        navegador.close()
        return produtos_promocao

def salvar_em_planilha(lista_produtos, pasta_alvo):
    if not lista_produtos:
        print("Nenhum produto em promoção para salvar.")
        return

    os.makedirs(pasta_alvo, exist_ok=True)
    df = pd.DataFrame(lista_produtos)

    caminho_excel = os.path.join(pasta_alvo, "promocoes_megaleste_1.xlsx")
    caminho_csv = os.path.join(pasta_alvo, "promocoes_megaleste_1.csv")

    # Salva os dados consolidados
    df.to_excel(caminho_excel, index=False, engine="openpyxl")
    df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")

    print(f"\n Total de ofertas extraídas das 57 páginas: {len(lista_produtos)}")
    print(f" Excel salvo em: {caminho_excel}")
    print(f"CSV salvo em: {caminho_csv}")

if __name__ == "__main__":
    dados = raspar_produtos_em_promocao()
    salvar_em_planilha(dados, PASTA_DESTINO)