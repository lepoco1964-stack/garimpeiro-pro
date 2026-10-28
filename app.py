import os
import json
import pandas as pd
from datetime import datetime, timedelta
from googleapiclient.discovery import build

# ⚙️ CONFIGURAÇÕES — AJUSTE AQUI
CHAVE_API = "COLE_SUA_CHAVE_AQUI"          # ← Coloque sua chave entre as aspas
NICHO = "música"                           # Palavra-chave do seu nicho
MAX_RESULTADOS_BUSCA = 50                  # Quantos canais buscar automaticamente
MAX_INSCRITOS = 5000                       # Apenas canais com até X inscritos
MIN_RELACAO_VIS_INS = 10                   # Mínimo: 10× visualizações sobre inscritos
IDADE_MAXIMA_DIAS = 180                    # Canais com menos de quantos dias

# 🎯 SEUS 10 CANAIS DE REFERÊNCIA — COLE AQUI
CANAIS_REFERENCIA = [
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx1",   # Canal 1 — cole o ID do canal
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx2",   # Canal 2
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx3",   # Canal 3
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx4",   # Canal 4
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx5",   # Canal 5
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx6",   # Canal 6
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx7",   # Canal 7
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx8",   # Canal 8
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx9",   # Canal 9
    "UCxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx10"   # Canal 10
]

youtube = build("youtube", "v3", developerKey=CHAVE_API)

def pegar_dados_canal_por_id(canal_id, origem="Busca automática"):
    """Pega dados completos de UM canal pelo ID"""
    try:
        stats = youtube.channels().list(
            id=canal_id, part="statistics,snippet"
        ).execute()["items"][0]
        
        nome = stats["snippet"]["title"]
        data_criacao = stats["snippet"]["publishedAt"]
        foto_url = stats["snippet"]["thumbnails"]["high"]["url"]
        pais = stats["snippet"].get("country", "Não informado")
        
        inscritos = int(stats["statistics"].get("subscriberCount", 0))
        total_visualizacoes = int(stats["statistics"].get("viewCount", 0))
        qtd_videos = int(stats["statistics"].get("videoCount", 0))
        
        media_por_video = total_visualizacoes / qtd_videos if qtd_videos > 0 else 0
        relacao = media_por_video / inscritos if inscritos > 0 else 0
        
        return {
            "nome": nome,
            "canal_id": canal_id,
            "inscritos": inscritos,
            "media_visualizacoes_video": round(media_por_video),
            "relacao_vid_ins": round(relacao, 2),
            "data_criacao": data_criacao[:10],
            "pais": pais,
            "foto_canal": foto_url,
            "url": f"https://youtube.com/channel/{canal_id}",
            "origem": origem
        }
    except Exception as e:
        print(f"⚠️ Erro ao buscar canal {canal_id}: {e}")
        return None

def buscar_canais_nicho(palavra_chave):
    """Busca canais novos no nicho com os filtros"""
    print(f"🔍 Buscando canais de '{palavra_chave}'...")
    canais_encontrados = []
    
    data_limite = (datetime.now() - timedelta(days=IDADE_MAXIMA_DIAS)).isoformat() + "Z"
    
    request = youtube.search().list(
        q=palavra_chave, type="channel", part="id,snippet",
        maxResults=MAX_RESULTADOS_BUSCA, publishedAfter=data_limite
    )
    response = request.execute()
    
    for item in response["items"]:
        canal_id = item["id"]["channelId"]
        dados = pegar_dados_canal_por_id(canal_id, origem="Busca automática")
        
        if dados and dados["inscritos"] > 0 and dados["inscritos"] <= MAX_INSCRITOS:
            if dados["relacao_vid_ins"] >= MIN_RELACAO_VIS_INS:
                canais_encontrados.append(dados)
    
    return canais_encontrados

# ▶️ EXECUTAR TUDO
if __name__ == "__main__":
    todos_canais = []
    
    # 1️⃣ Buscar canais novos
    canais_novos = buscar_canais_nicho(NICHO)
    todos_canais.extend(canais_novos)
    
    # 2️⃣ Adicionar seus canais de referência
    print(f"\n📋 Carregando seus {len(CANAIS_REFERENCIA)} canais de referência...")
    for canal_id in CANAIS_REFERENCIA:
        dados = pegar_dados_canal_por_id(canal_id, origem="✅ Referência")
        if dados:
            todos_canais.append(dados)
    
    # 💾 Salvar tudo no Excel
    df = pd.DataFrame(todos_canais)
    nome_arquivo = f"canais-garimpados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(nome_arquivo, index=False)
    
    # 📊 Resumo na tela
    print(f"\n✅ Concluído! {len(todos_canais)} canais processados")
    print(f"📂 Arquivo salvo: {nome_arquivo}")
    print("\n📋 Origem | Nome | Inscritos | Média Views | Relação")
    print("-" * 70)
    for _, row in df.iterrows():
        print(f"{row['origem']} | {row['nome'][:25]} | {row['inscritos']} | {row['media_visualizacoes_video']} | {row['relacao_vid_ins']}")
