import glob
from pathlib import Path
from diario_exonerados import extrair_texto, extrair_movimentacoes, gerar_planilha

def processar_lote():
    diretorio_pdfs = r"C:\Users\douglas.silva\.gemini\antigravity-ide\brain\b79e78e0-c1da-4b57-8d58-a978c3f973ec\.user_uploaded\*.pdf"
    arquivos_pdf = glob.glob(diretorio_pdfs)
    
    if not arquivos_pdf:
        print("Nenhum PDF encontrado na pasta de uploads.")
        return
        
    todas_movimentacoes = []
    
    for caminho in arquivos_pdf:
        print(f"Lendo: {Path(caminho).name}")
        texto = extrair_texto(Path(caminho))
        movimentacoes = extrair_movimentacoes(texto)
        print(f"  Encontradas {len(movimentacoes)} movimentações.")
        todas_movimentacoes.extend(movimentacoes)
        
    print(f"\nTotal acumulado: {len(todas_movimentacoes)} movimentações.")
    
    if todas_movimentacoes:
        print("Enviando para o Google Sheets...")
        gerar_planilha(todas_movimentacoes, None)
        print("Sucesso!")

if __name__ == "__main__":
    processar_lote()
