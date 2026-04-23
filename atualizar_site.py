import base64
import subprocess
import webbrowser
import sys
import os

print("============================================")
print("  Atualizador do Site - Buracos de Tatu")
print("============================================")
print()

# Recebe o caminho do arquivo (arrastar para cima do script ou digitar)
if len(sys.argv) > 1:
    arquivo = sys.argv[1].strip('"').strip("'")
else:
    arquivo = input("Digite ou arraste o caminho da planilha aqui: ").strip().strip('"').strip("'")

if not os.path.exists(arquivo):
    print()
    print("ERRO: Arquivo nao encontrado:", arquivo)
    input("Pressione Enter para fechar...")
    sys.exit(1)

print()
print("Convertendo planilha...")

with open(arquivo, "rb") as f:
    conteudo_b64 = base64.b64encode(f.read()).decode("ascii")

# Copia para area de transferencia
try:
    subprocess.run("clip", input=conteudo_b64.encode("ascii"), check=True)
    copiado = True
except Exception:
    copiado = False

print()
if copiado:
    print("PRONTO! Conteudo copiado para a area de transferencia.")
else:
    print("Nao foi possivel copiar automaticamente.")
    print("Copie manualmente o conteudo abaixo:")
    print()
    print(conteudo_b64)

print()
print("Agora:")
print("1. A pagina do GitHub vai abrir")
print("2. Clique em 'Run workflow' (botao cinza a direita)")
print("3. Cole o conteudo no campo 'planilha' (Ctrl+V)")
print("4. Clique em 'Run workflow' (botao verde)")
print()
input("Pressione Enter para abrir o GitHub...")

webbrowser.open("https://github.com/joaovtozoni/site_monitoramento_buracos_de_tatu/actions/workflows/gerar_site.yml")

input("Pressione Enter para fechar...")