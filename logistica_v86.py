# =============================================================
# compatível com Spyder E ambientes WEB/headless
# Pré-requisito: pip install googlemaps pandas openpyxl tqdm matplotlib numpy folium polyline
# =============================================================
import pandas as pd
import xml.etree.ElementTree as ET
import os, re, numpy as np
import unicodedata
import googlemaps
from datetime import datetime, date, timedelta
from tqdm import tqdm
import webbrowser as _webbrowser_mod  # importado uma vez; usado via salvar_e_abrir()

# ------------------------------------------------------------------
def salvar_e_abrir(caminho_html, silent=False):
    """
    Tenta abrir caminho_html no navegador padrão.
    Em ambientes sem display, silencia a exceção e registra o caminho.
    silent=True suprime até o print (útil para loops de muitos arquivos).
    """
    try:
        _webbrowser_mod.open('file:///' + str(caminho_html).replace('\\', '/'))
    except Exception:
        pass
    if not silent:
        print(f"   🌐 HTML salvo → {os.path.basename(str(caminho_html))}")

# =============================================================
# 1. CONFIGURAÇÕES
# =============================================================
# *** AJUSTE OS CAMINHOS ABAIXO PARA SUA MÁQUINA ***
DRIVE_PATH           = r'C:\directory'
XML_PATH             = r'C:\directory\notas_fiscais'
CACHE_FILE            = os.path.join(DRIVE_PATH, 'cache_logistica_google.xlsx')
DIRECTIONS_CACHE_FILE = os.path.join(DRIVE_PATH, 'cache_directions.pkl')
USO_API_FILE         = os.path.join(DRIVE_PATH, 'uso_api_google.xlsx')

# Velocidades usadas APENAS no fallback Haversine (quando API Google indisponível)
# A API Google retorna tempo real considerando tipo de via e trafego
VEL_URBANA_KMH   = 40   # distancias < 30 km (trechos urbanos)
VEL_RODOVIA_KMH  = 80   # distancias >= 30 km (rodovias)

API_KEY = 'GOOGLE_MAPS_API_KEY'
gmaps   = googlemaps.Client(
    key=API_KEY,
    timeout=20,                   # segundos por requisição (padrão 5 é curto no Colab)
    retry_over_query_limit=True   # retry automático se atingir cota por segundo
)

# --- Controle de custos API ---
CUSTO_POR_ELEMENTO_USD = 0.005
LIMITE_ELEMENTOS_DIA   = 2000
LIMITE_CUSTO_DIA_USD   = 10.00
ALERTA_ELEMENTOS       = 1600

# --- Engradados ---
TIPOS_ENGRADADOS = [
    {'nome': 'Half', 'dim': [1.2, 1.0, 0.7], 'vol': 0.84},
    {'nome': 'Full', 'dim': [2.4, 1.0, 0.7], 'vol': 1.68}
]

# Tipos de caminhão que NÃO permitem empilhamento de engradados
CAMINHOES_SEM_EMPILHAMENTO = {'fiorino', 'van'}

def permite_empilhamento(tipo_caminhao):
    """Retorna False para Fiorino e Van — sem empilhamento de engradados."""
    return str(tipo_caminhao).strip().lower() not in CAMINHOES_SEM_EMPILHAMENTO
ARMAZEM_SUZANO = {'lat': -23.598059, 'lon': -46.325861}

# --- Consolidação de rotas ---
RAIO_CONSOLIDACAO_KM = 75.0
# --- Clustering Geográfico (DBSCAN) ---
RAIO_DBSCAN_KM       = 50.0

# --- Entregas Interestaduais ---
# Distância mínima de Suzano (km) para classificar como interestadual
# quando a UF não puder ser lida da NF-e
DIST_INTERESTADUAL_KM = 300.0

# Regiões de consolidação: estados que podem compartilhar um caminhão
# Chave = nome da região, Valor = lista de UFs
REGIOES_CONSOLIDACAO = {
    'Sul':           ['PR', 'SC', 'RS'],
    'Sudeste':       ['RJ', 'MG', 'ES'],
    'Centro-Oeste':  ['GO', 'MT', 'MS', 'DF'],
    'Norte':         ['AM', 'PA', 'RO', 'AC', 'RR', 'AP', 'TO'],
    'Nordeste':      ['BA', 'SE', 'AL', 'PE', 'PB', 'RN', 'CE', 'PI', 'MA'],
}
# SP é sempre rota local — nunca entra no plano interestadual

# --- Deep Q-Network (DQN) — Seção 13F ---
RL_EPISODIOS         = 5000   # episódios de treino por execução (v81: era 2000)
RL_EPSILON_INI       = 1.0    # exploração inicial (100%)
RL_EPSILON_MIN       = 0.05   # exploração mínima (5%)
RL_EPSILON_DECAY     = 0.99940 # fator de decaimento por episódio (v85: calculado como 0.05^(1/5000) → ε=0.05 exatamente no ep 5000)
RL_GAMMA             = 0.95   # fator de desconto
RL_LR                = 1e-3   # learning rate (Adam)
RL_BATCH             = 64     # batch size do replay buffer
RL_BUFFER_MAX        = 2000   # capacidade máxima do replay buffer
RL_TARGET_UPDATE     = 20     # episódios entre atualizações da rede target (v81: era 50)
RL_PENALIDADE_ATR    = 500.0  # R$ por entrega fora do prazo (igual ao AG)
HORA_INICIO          = 8     # saída do armazém às 08:00
INTERVALO_DESPACHO   = 45    # minutos entre saídas consecutivas do armazém
PAUSA_ALMOCO_INI     = 12 * 60   # 12:00 em minutos desde meia-noite
PAUSA_ALMOCO_FIM     = 13 * 60   # 13:00
PAUSA_JANTA_INI      = 18 * 60   # 18:00
PAUSA_JANTA_FIM      = 19 * 60   # 19:00
TEMPO_DESCARGA_MIN   = 60    # minutos de descarga por parada (1 hora)

# Tempo de carregamento no armazém — varia com quantidade de engradados
TEMPO_CARGA_ATE3     = 15    # até 3 engradados: 15 min
TEMPO_CARGA_4A10     = 30    # 4 a 10 engradados: 30 min
TEMPO_CARGA_11MAIS   = 60    # 11 ou mais engradados: 60 min

def tempo_carregamento(qtd_engradados):
    """Retorna o tempo de carregamento em minutos com base na qtd de engradados."""
    if qtd_engradados <= 3:
        return TEMPO_CARGA_ATE3
    elif qtd_engradados <= 10:
        return TEMPO_CARGA_4A10
    else:
        return TEMPO_CARGA_11MAIS

STOPWORDS_ENDERECO = {
    'RUA', 'R', 'AV', 'AVENIDA', 'ALAMEDA', 'AL', 'ESTRADA', 'EST',
    'RODOVIA', 'ROD', 'TRAVESSA', 'TV', 'LARGO', 'LG', 'PRACA', 'PC',
    'VILA', 'JARDIM', 'JD', 'BAIRRO', 'BRO', 'CEP', 'BLOCO', 'BL',
    'APTO', 'AP', 'ANDAR', 'SALA', 'CJ', 'CONJ', 'LOTE', 'LT',
    'QUADRA', 'QD', 'SN', 'S/N', 'KM', 'DISTRITO', 'INDUSTRIAL',
    'DE', 'DA', 'DO', 'DAS', 'DOS', 'E', 'EM', 'NO', 'NA'
}


# =============================================================
# 2. CONTROLE DE CONSUMO DA API
# =============================================================

class ControladorAPI:
    """
    Rastreia consumo diário da Google Distance Matrix API.
    Persiste histórico em uso_api_google.xlsx no Drive.
    Bloqueia chamadas se limite de elementos ou custo for atingido.
    """
    def __init__(self, caminho):
        self.caminho        = caminho
        self.hoje           = str(date.today())
        self.elementos_hoje = 0
        self.chamadas_hoje  = 0
        self.custo_hoje_usd = 0.0
        self.bloqueado      = False
        self.historico      = []
        self._carregar()

    def _carregar(self):
        if os.path.exists(self.caminho):
            df = pd.read_excel(self.caminho)
            self.historico = df.to_dict('records')
            hoje_rows = [r for r in self.historico if str(r.get('data')) == self.hoje]
            if hoje_rows:
                self.elementos_hoje = sum(r.get('elementos', 0) for r in hoje_rows)
                self.chamadas_hoje  = sum(r.get('chamadas',  0) for r in hoje_rows)
                self.custo_hoje_usd = sum(r.get('custo_usd', 0.0) for r in hoje_rows)
        self._verificar_limites(silencioso=True)

    def _verificar_limites(self, silencioso=False):
        if (self.elementos_hoje >= LIMITE_ELEMENTOS_DIA or
                self.custo_hoje_usd >= LIMITE_CUSTO_DIA_USD):
            self.bloqueado = True
            if not silencioso:
                print(f"\n   🚫 LIMITE DIÁRIO ATINGIDO — API bloqueada.")
                print(f"      Elementos : {self.elementos_hoje}/{LIMITE_ELEMENTOS_DIA}")
                print(f"      Custo est.: US$ {self.custo_hoje_usd:.4f}/"
                      f"US$ {LIMITE_CUSTO_DIA_USD:.2f}")
                print(f"      Restante usará Haversine.")
        elif self.elementos_hoje >= ALERTA_ELEMENTOS and not silencioso:
            print(f"\n   ⚠️  Alerta: {self.elementos_hoje} elementos hoje — "
                  f"US$ {self.custo_hoje_usd:.4f}.")

    def pode_chamar(self, n_elem):
        if self.bloqueado:
            return False
        return ((self.elementos_hoje + n_elem) <= LIMITE_ELEMENTOS_DIA and
                (self.custo_hoje_usd + n_elem * CUSTO_POR_ELEMENTO_USD) <= LIMITE_CUSTO_DIA_USD)

    def registrar(self, n_chamadas, n_elementos, contexto=''):
        if self.bloqueado:
            return False
        custo = n_elementos * CUSTO_POR_ELEMENTO_USD
        self.chamadas_hoje  += n_chamadas
        self.elementos_hoje += n_elementos
        self.custo_hoje_usd += custo
        self.historico.append({
            'data':      self.hoje,
            'hora':      datetime.now().strftime('%H:%M:%S'),
            'chamadas':  n_chamadas,
            'elementos': n_elementos,
            'custo_usd': round(custo, 6),
            'contexto':  contexto,
            'acum_elem': self.elementos_hoje,
            'acum_usd':  round(self.custo_hoje_usd, 6)
        })
        pd.DataFrame(self.historico).to_excel(self.caminho, index=False)
        self._verificar_limites(silencioso=False)
        return True

    def resumo(self):
        print(f"\n   💰 Consumo API — {self.hoje}")
        print(f"      Chamadas  : {self.chamadas_hoje}")
        print(f"      Elementos : {self.elementos_hoje}/{LIMITE_ELEMENTOS_DIA}")
        print(f"      Custo est.: US$ {self.custo_hoje_usd:.4f}/"
              f"US$ {LIMITE_CUSTO_DIA_USD:.2f}")
        if self.bloqueado:
            print(f"      Status    : 🚫 BLOQUEADO")
        else:
            print(f"      Restantes : "
                  f"{LIMITE_ELEMENTOS_DIA - self.elementos_hoje} elem / "
                  f"US$ {LIMITE_CUSTO_DIA_USD - self.custo_hoje_usd:.4f}")


# =============================================================
# 3. DIAGNÓSTICO DA API GOOGLE
# =============================================================

def testar_api(gmaps_client, controlador):
    """
    Testa conectividade com a API com até 3 tentativas.
    Distingue Timeout (problema de rede) de erros de autenticação.
    """
    import socket
    if controlador.bloqueado:
        print("   🚫 Limite diário atingido — usando Haversine.")
        return False

    MAX_TENTATIVAS = 3
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            res    = gmaps_client.distance_matrix(
                origins=[(-23.598059, -46.325861)],
                destinations=[(-23.548943, -46.638819)],
                mode="driving"
            )
            status = res['rows'][0]['elements'][0]['status']
            controlador.registrar(1, 1, 'teste_conexao')
            if status == 'OK':
                km = res['rows'][0]['elements'][0]['distance']['value'] / 1000
                print(f"   ✅ API Google OK — rota de teste: {km:.1f} km")
                return True
            elif status == 'REQUEST_DENIED':
                print(f"   ❌ API KEY inválida ou Distance Matrix API não habilitada.")
                print(f"      Acesse: console.cloud.google.com → APIs → Distance Matrix API")
                return False
            else:
                print(f"   ⚠️  API status '{status}'.")
                return False

        except Exception as e:
            tipo_erro = type(e).__name__
            if 'Timeout' in tipo_erro or 'timeout' in str(e).lower():
                if tentativa < MAX_TENTATIVAS:
                    print(f"   ⏱️  Timeout na tentativa {tentativa}/{MAX_TENTATIVAS} "                          f"— aguardando 3s e tentando novamente...")
                    import time; time.sleep(3)
                    continue
                else:
                    print(f"   ⚠️  Timeout após {MAX_TENTATIVAS} tentativas.")
                    print(f"      A API será usada durante o processamento "                          f"(timeouts ocasionais são normais no Colab).")
                    print(f"      Se falhar novamente, o Haversine será usado como fallback.")
                    return True   # tenta usar a API — timeout pode ser temporário
            elif 'ApiError' in tipo_erro or '403' in str(e) or '400' in str(e):
                print(f"   ❌ Erro de autenticação: {repr(e)}")
                print(f"      Verifique: API_KEY correta e Distance Matrix API habilitada.")
                return False
            else:
                print(f"   ❌ Erro inesperado: {repr(e)}")
                return False

    return False


# =============================================================
# 4. FUNÇÕES AUXILIARES — PART NUMBER
# =============================================================

def normalizar_pn(pn):
    partes = str(pn).strip().upper().split('-')
    return '-'.join(partes[:2]) if len(partes) >= 3 else '-'.join(partes)


# =============================================================
# 5. FUNÇÕES AUXILIARES — ENDEREÇO
# =============================================================

def normalizar_endereco(texto):
    if not texto or str(texto).strip() in ('', 'nan'):
        return set()
    texto = unicodedata.normalize('NFD', str(texto).upper())
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto)
    return {t for t in texto.split()
            if t not in STOPWORDS_ENDERECO and len(t) >= 2}

def score_endereco(end_nf, end_cadastro):
    t_nf  = normalizar_endereco(end_nf)
    t_cad = normalizar_endereco(end_cadastro)
    if not t_nf or not t_cad:
        return 0.0
    intersec = t_nf & t_cad
    uniao    = t_nf | t_cad
    jaccard  = len(intersec) / len(uniao)
    nums_nf  = {t for t in t_nf  if t.isdigit()}
    nums_cad = {t for t in t_cad if t.isdigit()}
    bonus    = 0.2 if nums_nf and nums_cad and (nums_nf & nums_cad) else 0.0
    return min(jaccard + bonus, 1.0)


# =============================================================
# 6. BUSCA DE CLIENTE
# =============================================================

def buscar_cliente(nome_nf, end_nf, df_cli, score_minimo=0.10):
    col_nome = df_cli.iloc[:, 0].astype(str).str.strip().str.upper()
    nome_up  = nome_nf.strip().upper()

    def _melhor_match(candidatos):
        if candidatos.empty:
            return None, 0.0
        scores = candidatos.iloc[:, 1].astype(str).apply(
            lambda e: score_endereco(end_nf, e)
        )
        idx = scores.idxmax()
        return candidatos.loc[idx], float(scores[idx])

    for prefixo in [nome_up, nome_up[:15], nome_up[:8]]:
        cands = (df_cli[col_nome == prefixo] if prefixo == nome_up
                 else df_cli[col_nome.str.contains(prefixo, regex=False, na=False)])
        if not cands.empty:
            linha, sc = _melhor_match(cands)
            if linha is not None and sc >= score_minimo:
                _desc = TEMPO_DESCARGA_MIN
                if df_cli.shape[1] > 5:
                    try: _desc = int(float(str(linha.iloc[5]).replace(',','.'))) 
                    except: pass
                return (float(linha.iloc[2]), float(linha.iloc[3]),
                        str(linha.iloc[4]),   str(linha.iloc[0]), sc, _desc)

    palavras = [p for p in nome_up.split() if len(p) >= 4]
    if palavras:
        # Etapa 4a: primeira palavra significativa (resultado único)
        cands = df_cli[col_nome.str.contains(palavras[0], regex=False, na=False)]
        if len(cands) == 1:
            linha = cands.iloc[0]
            sc    = score_endereco(end_nf, str(linha.iloc[1]))
            print(f"   ⚠️  Match fraco '{palavras[0]}' (score:{sc:.2f}) — "
                  f"verifique '{nome_nf}'.")
            _desc = TEMPO_DESCARGA_MIN
            if df_cli.shape[1] > 5:
                try: _desc = int(float(str(linha.iloc[5]).replace(',','.')))
                except: pass
            return (float(linha.iloc[2]), float(linha.iloc[3]),
                    str(linha.iloc[4]),   str(linha.iloc[0]), sc, _desc)

        # Etapa 4b: qualquer candidato com score de endereço > 0
        # (último recurso — aceita mesmo com score baixo)
        if not cands.empty:
            linha, sc = _melhor_match(cands)
            if linha is not None and sc > 0:
                print(f"   ⚠️  Match por endereço '{palavras[0]}' (score:{sc:.2f}) — "
                      f"verifique '{nome_nf}'.")
                _desc = TEMPO_DESCARGA_MIN
                if df_cli.shape[1] > 5:
                    try: _desc = int(float(str(linha.iloc[5]).replace(',','.'))) 
                    except: pass
                return (float(linha.iloc[2]), float(linha.iloc[3]),
                        str(linha.iloc[4]),   str(linha.iloc[0]), sc, _desc)

    return None, None, None, None, 0.0, TEMPO_DESCARGA_MIN


# =============================================================
# 7. FUNÇÕES AUXILIARES — ENGRADADOS
# =============================================================

def cabe_em_engradado(dim_caixa, tipo_eng):
    """Verifica se uma caixa cabe no engradado (com rotação horizontal)."""
    c, l, a      = dim_caixa
    C, L, A      = tipo_eng['dim']
    normal       = (c <= C and l <= L and a <= A)
    rotacionado  = (l <= C and c <= L and a <= A)
    return normal or rotacionado

def separar_avulsos(itens):
    """
    Separa itens em dois grupos:
      - para_engradado : caixas que cabem em pelo menos um engradado
      - avulsos        : caixas que não cabem em nenhum engradado (vão direto no caminhão)
    """
    FULL = TIPOS_ENGRADADOS[1]   # maior engradado disponível
    para_engradado, avulsos = [], []
    for it in itens:
        if cabe_em_engradado(it['dim'], FULL):
            para_engradado.append(it)
        else:
            avulsos.append(it)
    return para_engradado, avulsos

def escolher_engradado(itens):
    """Versão original — usada no Nível 2 onde qtd=1 por item."""
    HALF           = TIPOS_ENGRADADOS[0]
    FULL           = TIPOS_ENGRADADOS[1]
    H_C, H_L, H_A = HALF['dim']
    for p in itens:
        c, l, a = p['dim']
        if not ((c <= H_C and l <= H_L and a <= H_A) or
                (l <= H_C and c <= H_L and a <= H_A)):
            return FULL
    if sum(np.prod(p['dim']) for p in itens) > HALF['vol']:
        return FULL
    return HALF


def motor_tetris(container_dim, itens, passo_x=0.1, sem_empilhamento=False):
    """
    Empacotamento 3-D com rotação horizontal (C↔L). Altura fixa.
    Salva posição (pos_x, pos_y, pos_z) de cada item para visualização 3D.
    Cada item na lista representa UMA peça física (qtd já expandida antes).
    sem_empilhamento=True: cada coluna só aceita 1 engradado (Fiorino, Van).
    """
    layout, rest        = [], []
    C_max, L_max, A_max = container_dim
    x, y, z, m_y       = 0.0, 0.0, 0.0, 0.0

    for it in sorted(itens, key=lambda i: np.prod(i['dim']), reverse=True):
        c, l, a     = it['dim']
        # Rotação horizontal: troca C e L, A sempre fixo
        orientacoes = [[c, l, a], [l, c, a]] if c != l else [[c, l, a]]
        alocado     = False

        # Tentativa 1: posição corrente (x, y, z)
        for ci, li, ai in orientacoes:
            # sem_empilhamento: só aceita z=0 (base do container)
            z_valido = (z == 0.0) if sem_empilhamento else True
            if (x + ci <= C_max and y + li <= L_max and
                    z + ai <= A_max and z_valido):
                it_c = it.copy()
                it_c['dim_f'] = (ci, li, ai)
                it_c['pos']   = (x, y, z)
                layout.append(it_c)
                z   += ai
                m_y  = max(m_y, li)
                alocado = True; break

        # Tentativa 2: nova pilha em Y (reinicia z)
        # sem_empilhamento: nova posição Y sempre com z=0
        if not alocado:
            z_new   = 0.0
            y_new   = y + m_y
            m_y_new = 0.0
            for ci, li, ai in orientacoes:
                if (x + ci <= C_max and y_new + li <= L_max and
                        ai <= A_max):
                    it_c = it.copy()
                    it_c['dim_f'] = (ci, li, ai)
                    it_c['pos']   = (x, y_new, z_new)
                    layout.append(it_c)
                    z   = ai if not sem_empilhamento else 0.0
                    y   = y_new
                    m_y = li
                    alocado = True; break

        # Tentativa 3: novo bloco em X (reinicia y e z)
        if not alocado:
            x_new = x + passo_x
            for ci, li, ai in orientacoes:
                if (x_new + ci <= C_max and li <= L_max and
                        ai <= A_max):
                    it_c = it.copy()
                    it_c['dim_f'] = (ci, li, ai)
                    it_c['pos']   = (x_new, 0.0, 0.0)
                    layout.append(it_c)
                    x   = x_new
                    y   = 0.0
                    z   = ai
                    m_y = li
                    alocado = True; break

        if not alocado:
            rest.append(it)

    return layout, rest


# =============================================================
# 8. FUNÇÕES AUXILIARES — CACHE E ROTAS
# =============================================================

def carregar_cache(caminho):
    if os.path.exists(caminho):
        df = pd.read_excel(caminho)
        if {'rota_id', 'km', 'minutos'}.issubset(df.columns):
            return dict(zip(df['rota_id'],
                            zip(df['km'].round(3), df['minutos'].round(1))))
    return {}

def salvar_cache(caminho, cache):
    pd.DataFrame([
        {'rota_id': k, 'km': round(v[0], 3), 'minutos': round(v[1], 1)}
        for k, v in cache.items()
    ]).to_excel(caminho, index=False)

def _rota_id(origem, destino):
    return (f"{round(origem[0], 4)},{round(origem[1], 4)}"
            f"|{round(destino[0], 4)},{round(destino[1], 4)}")

def _haversine(origem, destino, vel=None):
    """
    Distância em linha reta com velocidade média adaptativa:
      - Trecho urbano (< 30 km): 40 km/h
      - Rodovia (>= 30 km)     : 80 km/h
    Se vel for informado explicitamente, usa o valor fornecido.
    """
    R    = 6371.0
    dlat = np.radians(destino[0] - origem[0])
    dlon = np.radians(destino[1] - origem[1])
    a    = (np.sin(dlat / 2) ** 2
            + np.cos(np.radians(origem[0]))
            * np.cos(np.radians(destino[0]))
            * np.sin(dlon / 2) ** 2)
    dist = R * 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    if vel is None:
        vel = VEL_URBANA_KMH if dist < 30 else VEL_RODOVIA_KMH
    return round(dist, 3), round((dist / vel) * 60, 1)

def obter_rotas_lote(pares, cache, gmaps_client, api_ok, controlador, salvar=True):
    resultado, pares_novos = {}, []
    for orig, dest in pares:
        rid = _rota_id(orig, dest)
        if rid in cache:
            resultado[rid] = cache[rid]
        else:
            pares_novos.append((orig, dest))

    if not pares_novos:
        return resultado

    novas_entradas = {}
    BATCH = 10

    for i in range(0, len(pares_novos), BATCH):
        lote     = pares_novos[i:i + BATCH]
        origens  = list({p[0] for p in lote})
        destinos = list({p[1] for p in lote})
        n_elem   = len(origens) * len(destinos)
        api_usada = False

        if api_ok and controlador.pode_chamar(n_elem):
            try:
                res = gmaps_client.distance_matrix(
                    origins=origens, destinations=destinos, mode="driving"
                )
                for oi, orig in enumerate(origens):
                    for di, dest in enumerate(destinos):
                        elem = res['rows'][oi]['elements'][di]
                        rid  = _rota_id(orig, dest)
                        km, mins = (
                            (round(elem['distance']['value'] / 1000, 3),
                             round(elem['duration']['value'] / 60, 1))
                            if elem['status'] == 'OK'
                            else _haversine(orig, dest)
                        )
                        novas_entradas[rid] = resultado[rid] = (km, mins)
                controlador.registrar(1, n_elem, f'lote_{i//BATCH+1}')
                api_usada = True
            except Exception as e:
                print(f"   ⚠️  Erro API lote {i//BATCH+1}: {repr(e)}")
        elif api_ok and not controlador.bloqueado:
            print(f"   🚫 Lote cancelado — limite diário seria ultrapassado. "
                  f"Usando Haversine.")
            controlador.bloqueado = True

        if not api_usada:
            for orig, dest in lote:
                rid = _rota_id(orig, dest)
                novas_entradas[rid] = resultado[rid] = _haversine(orig, dest)

    if novas_entradas:
        cache.update(novas_entradas)
        if salvar:
            salvar_cache(CACHE_FILE, cache)
        print(f"   ✅ {len(novas_entradas)} rota(s) nova(s) no cache.")

    return resultado


# =============================================================
# 9. NÍVEL 2 — FUNÇÕES DE CONSOLIDAÇÃO
# =============================================================

def ajustar_pausas(partida_dt, chegada_dt):
    """
    Ajusta o horário de chegada considerando pausas obrigatórias do motorista.
    Se o trajeto atravessar o horário de almoço (12-13h) ou janta (18-19h),
    adiciona 1h à chegada.
    Suporta viagens que cruzam meia-noite (usa minutos absolutos desde 00:00).
    """
    # Converte para minutos desde 00:00 do dia base
    def to_min(dt):
        base = dt.replace(hour=0, minute=0, second=0, microsecond=0)
        return int((dt - base).total_seconds() / 60) + (
            1440 if dt.date() > partida_dt.date() else 0
        )

    p_min = to_min(partida_dt)
    c_min = to_min(chegada_dt)
    atraso = 0

    for ini, fim in [(PAUSA_ALMOCO_INI, PAUSA_ALMOCO_FIM),
                     (PAUSA_JANTA_INI,  PAUSA_JANTA_FIM)]:
        # Pausa dentro do trajeto: partida antes do fim da pausa
        #                          chegada (sem atraso) depois do início
        if p_min < fim and (c_min + atraso) > ini:
            # Quanto do trajeto cai dentro da pausa
            inicio_pausa = max(p_min, ini)
            fim_pausa    = min(c_min + atraso, fim)
            if fim_pausa > inicio_pausa:
                atraso += fim - ini   # adiciona 1h inteira

    return chegada_dt + timedelta(minutes=atraso)


# Data base de referência para cálculo de minutos absolutos
# (dia de saída do armazém = hoje às 00:00)
_DIA_BASE = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

def lim_min(limite_str):
    """
    Converte limite de entrega em minutos desde 00:00 do dia de saída.
    Formatos aceitos:
      'HH:MM'          → mesmo dia  (ex: '18:00' = 1080 min)
      'HH:MM:SS'       → mesmo dia  (ex: '17:00:00' = 1020 min — Excel salva assim)
      'HH:MM+1'        → dia seguinte (ex: '02:00+1' = 1560 min)
    Padrão se inválido: 18:00 do mesmo dia = 1080 min.
    """
    try:
        s = str(limite_str).strip()
        dia_offset = 1 if s.endswith('+1') else 0
        s = s.replace('+1', '')
        partes = s.split(':')
        h = int(partes[0])
        m = int(partes[1]) if len(partes) > 1 else 0
        # ignora segundos se presentes (partes[2])
        return h * 60 + m + dia_offset * 1440
    except Exception:
        return 18 * 60

def dt_min(dt):
    """Minutos desde 00:00 do dia base (suporta entregas no dia seguinte)."""
    delta = dt - _DIA_BASE
    return int(delta.total_seconds() / 60)


def construir_grupos_destino(eng_final):
    """
    Agrupa engradados pelo destino (cli).
    Múltiplas NFs para o mesmo destino formam um único grupo.
    O limite de entrega do grupo é o mais restritivo entre as NFs.
    """
    grupos = {}
    for eng in eng_final:
        cli = eng['cli']
        if cli not in grupos:
            grupos[cli] = {
                'cli':      cli,
                'lat':      eng['lat'],
                'lon':      eng['lon'],
                'limite':   eng['limite'],
                'descarga': eng.get('descarga', TEMPO_DESCARGA_MIN),
                'nfs':      set(),
                'engs':     []
            }
        grupos[cli]['nfs'].add(eng['nf'])
        if lim_min(eng['limite']) < lim_min(grupos[cli]['limite']):
            grupos[cli]['limite'] = eng['limite']
        # Usa o maior tempo de descarga entre NFs do mesmo cliente
        grupos[cli]['descarga'] = max(
            grupos[cli]['descarga'],
            eng.get('descarga', TEMPO_DESCARGA_MIN)
        )
        grupos[cli]['engs'].append(eng)
    return list(grupos.values())


def consolidar_por_proximidade(grupos):
    """
    Agrupa destinos próximos em clusters para compartilhar viagem.
    Algoritmo greedy: ordena pelo limite mais urgente, expande por raio.
    Garante que todos os destinos do cluster sejam temporalmente compatíveis
    (o mais urgente define o limite do cluster inteiro).
    """
    pendentes = sorted(grupos, key=lambda g: lim_min(g['limite']))
    alocados  = set()
    clusters  = []

    for i, ancora in enumerate(pendentes):
        if i in alocados:
            continue
        cluster   = [ancora]
        alocados.add(i)
        lim_cluster = lim_min(ancora['limite'])  # limite mais restritivo do cluster

        for j, cand in enumerate(pendentes):
            if j in alocados:
                continue
            dist = _haversine(
                (ancora['lat'], ancora['lon']),
                (cand['lat'],   cand['lon'])
            )[0]
            # Só consolida se dentro do raio E se o candidato
            # tem limite compatível com o cluster (não mais restritivo que a âncora)
            if dist <= RAIO_CONSOLIDACAO_KM and lim_min(cand['limite']) >= lim_cluster:
                cluster.append(cand)
                alocados.add(j)

        clusters.append(cluster)

    return clusters


def simular_viagem(paradas_ord, cache, gmaps_client, api_ok, controlador,
                   h_inicio=None):
    """
    Simula a viagem na ordem dada e verifica se todos os clientes
    são atendidos dentro de seus limites de horário.
    h_inicio: datetime de saída do armazém (default: HORA_INICIO do dia)
    Retorna (rotas_dict, viavel, chegadas_dict).
    """
    h_ref  = h_inicio if h_inicio else datetime.today().replace(
        hour=HORA_INICIO, minute=0, second=0, microsecond=0
    )
    p_ref  = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
    pares  = []
    p_tmp  = p_ref

    for g in paradas_ord:
        dest = (g['lat'], g['lon'])
        pares.append((p_tmp, dest))
        p_tmp = dest

    rotas    = obter_rotas_lote(pares, cache, gmaps_client, api_ok, controlador)
    viavel   = True
    chegadas = {}
    h_cur    = h_ref
    p_cur    = p_ref

    for g in paradas_ord:
        dest     = (g['lat'], g['lon'])
        rid      = _rota_id(p_cur, dest)
        _, mins  = rotas.get(rid, _haversine(p_cur, dest))
        chegada_bruta = h_cur + timedelta(minutes=mins)
        chegada       = ajustar_pausas(h_cur, chegada_bruta)   # pausa almoço/janta
        chegadas[g['cli']] = chegada
        if dt_min(chegada) > lim_min(g['limite']):
            viavel = False
        desc_g = g.get('descarga', TEMPO_DESCARGA_MIN)
        h_cur  = chegada + timedelta(minutes=desc_g)
        p_cur  = dest

    return rotas, viavel, chegadas


def empacotar_em_caminhao(grupos_cluster, truck_list):
    """
    Converte engradados e caixas avulsas do cluster em peças para o tetris.

    Estratégia:
      1. Testa caminhões do MENOR para o MAIOR.
      2. Para cada caminhão, verifica primeiro se TODAS as peças cabem
         geometricamente (verificação de volume e dimensão antes do tetris).
      3. Só aceita um caminhão se ele comportar TODAS as peças sem sobras.
      4. Se nenhum caminhão comportar tudo, usa o maior disponível e
         retorna as sobras para reprocessamento pelo loop principal.

    Retorna (truck, alocados, sobram).
    """
    pecas = []
    for g in grupos_cluster:
        for eng in g['engs']:
            pecas.append({**eng})

    if not pecas:
        return None, [], []

    pecas_normais = [p for p in pecas if p.get('tipo') != 'Avulso']
    pecas_avulsas = [p for p in pecas if p.get('tipo') == 'Avulso']
    todas         = pecas_normais + pecas_avulsas

    def _cabe_geometricamente(peca, truck):
        """Verifica se uma peça cabe no caminhão em alguma orientação horizontal."""
        c, l, a = peca['dim']
        C, L, A = truck['dim']
        return ((c <= C and l <= L and a <= A) or
                (l <= C and c <= L and a <= A))

    def _vol_total_pecas(pecas_lista):
        return sum(p['dim'][0]*p['dim'][1]*p['dim'][2] for p in pecas_lista)

    vol_total = _vol_total_pecas(todas)
    maior     = max(todas, key=lambda p: np.prod(p['dim']), default=None)

    # Filtra caminhões que passam nos dois critérios:
    #   1. Volume do caminhão >= volume total das peças
    #   2. Maior peça cabe geometricamente (dimensões)
    candidatos = []
    for truck in truck_list:
        vol_truck = truck['vol']   # usa vol pré-calculado (Length×Width×Height do Excel)
        if vol_truck < vol_total:
            continue   # volume insuficiente — descarta
        if maior and not _cabe_geometricamente(maior, truck):
            continue   # maior peça não cabe geometricamente — descarta
        candidatos.append((truck, vol_truck))
    
    # Debug: mostra por que cada caminhão foi aceito/rejeitado
    # (remova após confirmar funcionamento)
    for truck in truck_list:
        vt = truck['vol']
        m_cabe = not maior or _cabe_geometricamente(maior, truck)
        status = '✅' if (vt >= vol_total and m_cabe) else '❌'
        motivo = '' if status == '✅' else (
            f'vol {vt:.2f}<{vol_total:.2f}' if vt < vol_total else 'dim insuf'
        )
        # print(f"      {status} {truck['tipo']}: vol={vt:.2f} {motivo}")

    if not candidatos:
        # Nenhum caminhão comporta tudo — usa o maior disponível com sobras
        truck_maior  = truck_list[-1]
        passo        = max(0.1, min(p['dim'][0] for p in todas))
        sem_empilh   = not permite_empilhamento(truck_maior['tipo'])
        alocados, sobram = motor_tetris(truck_maior['dim'], todas,
                                        passo_x=passo, sem_empilhamento=sem_empilh)
        return truck_maior, alocados, sobram

    # Testa do menor para o maior caminhão candidato
    for truck, vol_truck in candidatos:
        passo        = max(0.1, min(p['dim'][0] for p in todas))
        sem_empilh   = not permite_empilhamento(truck['tipo'])
        alocados, sobram = motor_tetris(truck['dim'], todas,
                                        passo_x=passo, sem_empilhamento=sem_empilh)

        if not alocados:
            continue

        if not sobram:
            # Tudo coube — retorna o menor caminhão suficiente
            return truck, alocados, sobram

    # Candidatos não conseguiram alocar tudo pelo tetris (geometria interna)
    # Usa o maior candidato e retorna sobras
    truck_maior, _ = candidatos[-1]
    passo          = max(0.1, min(p['dim'][0] for p in todas))
    sem_empilh     = not permite_empilhamento(truck_maior['tipo'])
    alocados, sobram = motor_tetris(truck_maior['dim'], todas,
                                    passo_x=passo, sem_empilhamento=sem_empilh)
    return truck_maior, alocados, sobram


# =============================================================
# 10. CARREGAMENTO DAS BASES
# =============================================================
print("\n📂 Carregando bases de dados...")

controlador_api = ControladorAPI(USO_API_FILE)
print(f"   📊 API hoje: {controlador_api.elementos_hoje} elem / "
      f"US$ {controlador_api.custo_hoje_usd:.4f}")
if controlador_api.bloqueado:
    print(f"   🚫 Limite diário atingido — usará Haversine.")

mem_cache = carregar_cache(CACHE_FILE)
print(f"   💾 Rotas em cache: {len(mem_cache)}")

# Carrega cache de polylines (Directions API) do disco
import pickle
_cache_directions = {}
if os.path.exists(DIRECTIONS_CACHE_FILE):
    try:
        with open(DIRECTIONS_CACHE_FILE, 'rb') as _f:
            _cache_directions = pickle.load(_f)
        print(f"   💾 Polylines em cache: {len(_cache_directions)} rotas")
    except Exception:
        _cache_directions = {}

df_dim    = pd.read_excel(os.path.join(DRIVE_PATH, 'dimensoes_produtos.xlsx'))
df_cli    = pd.read_excel(os.path.join(DRIVE_PATH, 'Cadastro Clientes.xlsx'))
df_trucks = pd.read_excel(os.path.join(DRIVE_PATH, 'Truck Types.xlsx'), dtype=str)

if df_cli.shape[1] < 5:
    raise ValueError("Cadastro Clientes: A=Nome|B=Endereço|C=Lat|D=Lon|E=Limite")

end_ok = df_cli.iloc[:, 1].notna().sum()
if end_ok == 0:
    print("   ⚠️  Coluna B (endereço) vazia — desambiguação de filiais prejudicada.")

# Coluna "Horario Negociado" — busca pelo nome do header (robusto a reordenação)
# Aceita qualquer coluna cujo header contenha "negociado" ou "horario neg"
# Formato: HH:MM (mesmo dia) ou HH:MM+1 (dia seguinte)
limite_negociado = {}  # {nome_cliente_upper: 'HH:MM' ou 'HH:MM+1'}

_col_neg = None
for _col in df_cli.columns:
    _col_str = str(_col).strip().lower()
    if 'negociado' in _col_str or 'horario neg' in _col_str or 'hor neg' in _col_str:
        _col_neg = _col
        break

# Fallback: tenta por índice (G=6, H=7) se header não encontrado
if _col_neg is None:
    for _idx in [6, 7, 5]:   # tenta G, H, F
        if df_cli.shape[1] > _idx:
            _amostra = df_cli.iloc[:, _idx].dropna().astype(str)
            # Verifica se a coluna parece conter horários (HH:MM ou HH:MM+1)
            _tem_hora = _amostra.str.match(r'^\d{1,2}:\d{2}').sum()
            if _tem_hora > 0:
                _col_neg = df_cli.columns[_idx]
                print(f"   ℹ️  Coluna de horário negociado detectada por conteúdo: "
                      f"coluna {_idx+1} ('{_col_neg}')")
                break

if _col_neg is not None:
    for _, row in df_cli.iterrows():
        nome_cli = str(row.iloc[0]).strip()
        val      = str(row[_col_neg]).strip()
        if val and val.lower() not in ('nan', 'none', ''):
            limite_negociado[nome_cli.upper()] = val
    n_neg = len(limite_negociado)
    if n_neg > 0:
        print(f"   🤝 Limites negociados carregados: {n_neg} cliente(s) "
              f"(coluna '{_col_neg}')")
    else:
        print(f"   ℹ️  Coluna '{_col_neg}' encontrada mas sem valores — "
              f"usando limite padrão.")
else:
    print("   ℹ️  Coluna de horário negociado não encontrada — usando limite padrão.")
    print("         Adicione coluna com header 'Horario Negociado' no cadastro.")

def limite_otimizado_cliente(nome_cli):
    """
    Retorna o limite de entrega negociado para o cliente.
    Busca pelo nome exato (case-insensitive) e por prefixo de 20 chars.
    Fallback: LIMITE_OTIMIZADO (02:00+1).
    """
    chave = str(nome_cli).strip().upper()
    if chave in limite_negociado:
        return limite_negociado[chave]
    # Tenta prefixo de 20 chars (cobre nomes truncados)
    for k, v in limite_negociado.items():
        if chave[:20] == k[:20]:
            return v
    return LIMITE_OTIMIZADO

df_dim['pn_norm'] = df_dim.iloc[:, 5].astype(str).apply(normalizar_pn)

def _local_carga(row):
    """Lê coluna H (Local de Carregamento) do Truck Types. Default: Doca."""
    if df_trucks.shape[1] > 7:
        val = str(row.iloc[7]).strip().lower()
        if 'pateo' in val or 'páteo' in val or 'patio' in val or 'pátio' in val:
            return 'Pateo'
    return 'Doca'

truck_list = sorted([
    {'tipo':  str(r.iloc[1]).strip(),
     'dim':   [float(r.iloc[2]), float(r.iloc[3]), float(r.iloc[4])],
     'vol':   float(r.iloc[2]) * float(r.iloc[3]) * float(r.iloc[4]),
     'local': _local_carga(r)}
    for _, r in df_trucks.iterrows()
], key=lambda x: x['vol'])

_docas  = [t['tipo'] for t in truck_list if t['local'] == 'Doca']
_pateos = [t['tipo'] for t in truck_list if t['local'] == 'Pateo']
print(f"   🏭 Doca : {', '.join(_docas)}")
print(f"   🚗 Pateo: {', '.join(_pateos)}")

print(f"   ✅ Clientes:{len(df_cli)} | Produtos:{len(df_dim)} | Caminhões:{len(truck_list)}")

print("\n   === Nomes no Cadastro Clientes.xlsx ===")
for i, nome in enumerate(df_cli.iloc[:, 0].astype(str)):
    print(f"   [{i:03d}] '{nome}'")

print("\n🔌 Testando API Google...")
API_DISPONIVEL = testar_api(gmaps, controlador_api)


# =============================================================
# 11. NÍVEL 1 — MONTAGEM DOS ENGRADADOS
# =============================================================
print("\n📦 Nível 1: Processando Engradados...")

# -----------------------------------------------------------------
# PRÉ-COMPUTAÇÕES (feitas uma vez, fora do loop — ganho significativo)
# -----------------------------------------------------------------

# 1. Índice de PN: dict {pn_norm: (dim_c, dim_l, dim_a, qtd_lote)} para busca O(1)
#    Coluna F (índice 5) = Part Number
#    Coluna H (índice 7) = Quantidade do lote por caixa
#    Coluna I (índice 8) = Comprimento da caixa (m)
#    Coluna J (índice 9) = Largura da caixa (m)
#    Coluna K (índice 10)= Altura da caixa (m)
indice_pn = {}
for _, row in df_dim.iterrows():
    pn_key = str(row['pn_norm']).strip().upper()
    if pn_key and pn_key not in indice_pn:
        try:
            qtd_lote = int(float(row.iloc[7]))   # col H — qtd por caixa
            if qtd_lote <= 0:
                qtd_lote = 1
            indice_pn[pn_key] = (
                float(row.iloc[8]),   # col I — comprimento caixa (m)
                float(row.iloc[9]),   # col J — largura caixa (m)
                float(row.iloc[10]),  # col K — altura caixa (m)
                qtd_lote              # col H — peças por caixa
            )
        except Exception:
            pass
print(f"   🔍 Índice de PNs construído: {len(indice_pn)} entradas")

# 2. Pré-normaliza coluna de nome do cadastro (evita recalcular a cada NF)
df_cli['_nome_upper'] = df_cli.iloc[:, 0].astype(str).str.strip().str.upper()

# 3. Cache de clientes já resolvidos: {nome_nf+end_nf: resultado}
#    Clientes que aparecem em múltiplas NFs são resolvidos apenas uma vez
cache_clientes = {}

import time
t_inicio_n1  = time.time()
t_xml        = 0.0
t_pn         = 0.0
t_cliente    = 0.0
t_tetris     = 0.0
nfs_processadas = 0

eng_final            = []
rel_eng              = []
clientes_nao_achados = {}
pns_nao_achados      = {}   # {descricao: qtd_total}
id_gen               = 1

for f in tqdm(sorted([x for x in os.listdir(XML_PATH) if x.endswith('.xml')])):

    # --- Parse XML com tolerância a encoding incorreto ---
    _t = time.time()
    try:
        root = ET.parse(os.path.join(XML_PATH, f)).getroot()
    except ET.ParseError:
        # Arquivo com encoding Latin-1 declarado como UTF-8 (ex: símbolo °)
        # Lê como bytes, decodifica com Latin-1 e sanitiza
        with open(os.path.join(XML_PATH, f), 'rb') as _fxml:
            _raw = _fxml.read()
        try:
            _txt = _raw.decode('utf-8')
        except UnicodeDecodeError:
            _txt = _raw.decode('latin-1')
        # Remove bytes inválidos para XML substituindo por '?'
        _txt = ''.join(c if ord(c) < 0xD800 or ord(c) > 0xDFFF else '?' for c in _txt)

        try:
            root = ET.fromstring(_txt)
        except ET.ParseError as e2:
            print(f"\n   ❌ NF {f}: XML inválido mesmo após sanitização — {e2}. Ignorada.")
            continue
    ns   = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
    nf   = root.find('.//ns:nNF',           ns).text
    cli  = root.find('.//ns:dest/ns:xNome', ns).text

    def _txt(xpath):
        node = root.find(xpath, ns)
        return node.text.strip() if node is not None and node.text else ''

    end_nf = ' '.join(filter(None, [
        _txt('.//ns:dest/ns:enderDest/ns:xLgr'),
        _txt('.//ns:dest/ns:enderDest/ns:nro'),
        _txt('.//ns:dest/ns:enderDest/ns:xBairro')
    ]))
    uf_dest = _txt('.//ns:dest/ns:enderDest/ns:UF').strip().upper()  # ex: 'SP', 'PR', 'RJ'
    t_xml += time.time() - _t

    # --- Coleta caixas via índice O(1) ---
    # Lógica correta:
    #   qtd_NF   = quantidade de peças na nota fiscal
    #   qtd_lote = peças que cabem em uma caixa (col H do Excel)
    #   n_caixas = ceil(qtd_NF / qtd_lote)
    # Cada caixa é uma unidade para o tetris, com dimensões I×J×K.
    # Um PN com 236.504 peças e lote de 100 → 2.366 caixas idênticas.
    import math
    _t = time.time()
    itens_nf = []
    for det in root.findall('.//ns:det', ns):
        pn_nf   = str(det.find('.//ns:prod/ns:cProd', ns).text).strip().upper()
        pn_norm = normalizar_pn(pn_nf)
        qtd_nf  = int(float(det.find('.//ns:prod/ns:qCom', ns).text))
        info    = indice_pn.get(pn_norm)
        if info:
            dim_c, dim_l, dim_a, qtd_lote = info
            d_p      = [dim_c, dim_l, dim_a]
            vol_cx   = round(dim_c * dim_l * dim_a, 6)
            n_caixas = math.ceil(qtd_nf / qtd_lote)   # nº de caixas para esta linha da NF
            for _ in range(n_caixas):
                itens_nf.append({
                    'dim':      d_p,
                    'vol':      vol_cx,
                    'pn':       pn_nf,
                    'qtd_nf':   qtd_nf,
                    'qtd_lote': qtd_lote,
                    'n_caixas': n_caixas
                })
        else:
            chave_pn = f"{pn_nf} (norm: {pn_norm})"
            pns_nao_achados[chave_pn] = pns_nao_achados.get(chave_pn, 0) + qtd_nf
    t_pn += time.time() - _t

    if not itens_nf:
        continue

    # --- Busca cliente (com cache de resultado) ---
    _t = time.time()
    chave_cli = f"{cli}||{end_nf}"
    if chave_cli in cache_clientes:
        lat, lon, lim, nome_cadastro, score = cache_clientes[chave_cli]
    else:
        lat, lon, lim, nome_cadastro, score, descarga_cli = buscar_cliente(cli, end_nf, df_cli)
        cache_clientes[chave_cli] = (lat, lon, lim, nome_cadastro, score)
    t_cliente += time.time() - _t

    if lat is None:
        clientes_nao_achados[cli] = end_nf
        print(f"   ❌ '{cli}' | '{end_nf}' | NF {nf} — ignorada.")
        continue

    tokens_end       = sorted(normalizar_endereco(end_nf) - {'S', 'N'})[:3]
    nome_agrupamento = (f"{nome_cadastro} [{' '.join(tokens_end)}]"
                        if tokens_end else nome_cadastro)

    # --- Separa caixas avulsas (não cabem em nenhum engradado) ---
    _t = time.time()
    itens_nf, avulsos_nf = separar_avulsos(itens_nf)

    if avulsos_nf:
        print(f"   📦 NF {nf}: {len(avulsos_nf)} caixa(s) avulsa(s) "
              f"(não cabem em engradado) → carga direta no caminhão.")

    # --- Tetris: passo_x=0.2 ---
    while itens_nf:
        escolha       = escolher_engradado(itens_nf)
        lay, itens_nf = motor_tetris(escolha['dim'], itens_nf, passo_x=0.2)
        if not lay:
            print(f"   ⚠️  Itens residuais NF {nf} sem encaixe — descartados.")
            break

        # Volume real = soma dos volumes individuais de cada caixa alocada
        # dim_f = dimensão da caixa DENTRO do engradado (após rotação)
        # O volume de uma caixa nunca pode exceder o volume do engradado inteiro
        vol_pecas = 0.0
        for p in lay:
            if 'dim_f' in p:
                vc = p['dim_f'][0] * p['dim_f'][1] * p['dim_f'][2]
            else:
                vc = p['dim'][0] * p['dim'][1] * p['dim'][2]
            # Cada caixa individual não pode ter volume > engradado
            vol_pecas += min(vc, escolha['vol'])
        vol_pecas  = float(min(vol_pecas, escolha['vol']))   # cap no total
        eficiencia = round((vol_pecas / escolha['vol']) * 100, 2)

        if eficiencia > 100.0:
            print(f"   ⚠️  Sanidade: eng {id_gen:03d} efic={eficiencia}% "
                  f"vol_cx={vol_pecas:.4f} vol_eng={escolha['vol']:.4f} "
                  f"n_cx={len(lay)} — caixas maiores que engradado?")
            eficiencia = 100.0

        id_box = f"ENG_{id_gen:03d}"

        # Consolida info de caixas por PN para o relatório
        info_caixas = {}
        for p in lay:
            pn = p['pn']
            if pn not in info_caixas:
                info_caixas[pn] = {
                    'n_caixas': 0,
                    'qtd_lote': p.get('qtd_lote', 1),
                    'qtd_nf':   p.get('qtd_nf', 0)
                }
            info_caixas[pn]['n_caixas'] += 1

        resumo_pns = "; ".join(
            f"{pn}({v['n_caixas']}cx × {v['qtd_lote']}pç = {v['n_caixas']*v['qtd_lote']}pç)"
            for pn, v in sorted(info_caixas.items())
        )

        rel_eng.append({
            'ID Sequencia':             id_box,
            'Nota Fiscal':              nf,
            'Cliente NF':               cli,
            'Cliente Cadastro':         nome_cadastro,
            'Endereço NF':              end_nf,
            'Score Endereço':           round(score, 2),
            'Tipo Engradado':           escolha['nome'],
            'Part Numbers':             ", ".join(sorted(info_caixas.keys())),
            'Detalhamento Caixas':      resumo_pns,
            'Qtd Caixas Engradado':     len(lay),
            'Volume Caixas (m³)':       round(vol_pecas, 4),
            'Volume Engradado (m³)':    round(escolha['vol'], 4),
            'Eficiência Volumétrica %': eficiencia
        })

        eng_final.append({
            'id':     id_box,
            'dim':    escolha['dim'],
            'vol':    escolha['vol'],
            'tipo':    escolha['nome'],
            'cli':     nome_agrupamento,
            'nf':      nf,
            'lat':     lat,
            'lon':     lon,
            'limite':  lim,
            'descarga': descarga_cli,
            'uf':      uf_dest,
        })
        id_gen += 1

    # Registra caixas avulsas — cada uma é uma "unidade" no planejamento
    for av in avulsos_nf:
        vol_av = round(av['dim'][0] * av['dim'][1] * av['dim'][2], 4)
        id_box = f"ENG_{id_gen:03d}"
        rel_eng.append({
            'ID Sequencia':             id_box,
            'Nota Fiscal':              nf,
            'Cliente NF':               cli,
            'Cliente Cadastro':         nome_cadastro,
            'Endereço NF':              end_nf,
            'Score Endereço':           round(score, 2),
            'Tipo Engradado':           'Avulso',
            'Part Numbers':             av['pn'],
            'Detalhamento Caixas':      f"{av['pn']}(1cx × {av.get('qtd_lote',1)}pç)",
            'Qtd Caixas Engradado':     1,
            'Volume Caixas (m³)':       vol_av,
            'Volume Engradado (m³)':    vol_av,
            'Eficiência Volumétrica %': 100.0
        })
        eng_final.append({
            'id':      id_box,
            'dim':     av['dim'],
            'vol':     vol_av,
            'tipo':    'Avulso',
            'cli':     nome_agrupamento,
            'nf':      nf,
            'lat':     lat,
            'lon':     lon,
            'limite':  lim,
            'descarga': descarga_cli,
            'uf':      uf_dest,
        })
        id_gen += 1

    t_tetris += time.time() - _t
    nfs_processadas += 1

# --- Relatório de performance ---
t_total_n1 = time.time() - t_inicio_n1
print(f"\n   ⏱️  Nível 1 concluído em {t_total_n1:.1f}s "
      f"({nfs_processadas} NFs processadas)")
print(f"   {'Etapa':<12} {'Total(s)':>9} {'Média/NF(ms)':>14} {'%':>6}")
print(f"   {'-'*45}")
for etapa, t in [('XML parse', t_xml), ('Busca PN',  t_pn),
                 ('Cliente',   t_cliente), ('Tetris', t_tetris)]:
    media = (t / nfs_processadas * 1000) if nfs_processadas else 0
    pct   = (t / t_total_n1 * 100) if t_total_n1 else 0
    print(f"   {etapa:<12} {t:>9.2f} {media:>14.1f} {pct:>5.1f}%")
print(f"   {'TOTAL':<12} {t_total_n1:>9.2f}")

pd.DataFrame(rel_eng).to_excel(os.path.join(DRIVE_PATH, 'Engradados.xlsx'), index=False)
print(f"\n   ✅ {id_gen-1} engradado(s) → Engradados.xlsx")

if pns_nao_achados:
    print(f"\n   ⚠️  PNs sem dimensão ({len(pns_nao_achados)}):")
    print(f"   {'PN (normalizado)':<45} {'Qtd NF':>8}")
    print(f"   {'-'*55}")
    for pn, qtd in sorted(pns_nao_achados.items()):
        print(f"       • {pn:<43} {qtd:>8,}")

if clientes_nao_achados:
    print(f"\n   ❌ Clientes não encontrados:")
    for nome, end in sorted(clientes_nao_achados.items()):
        print(f"       '{nome}' | '{end}'")


# =============================================================
# 11B. CLASSIFICAÇÃO INTERESTADUAL
# Separa engradados de outros estados antes do planejamento diário.
# Critério: UF da NF-e (principal) ou distância de Suzano (fallback).
# =============================================================

def _uf_do_engradado(eng):
    """Extrai UF do engradado. Tenta campo 'uf', depois distância."""
    uf = eng.get('uf', '').strip().upper()
    if uf and len(uf) == 2:
        return uf
    # Fallback: distância haversine de Suzano
    dist, _ = _haversine(
        (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']),
        (eng.get('lat', ARMAZEM_SUZANO['lat']),
         eng.get('lon', ARMAZEM_SUZANO['lon']))
    )
    return 'SP' if dist <= DIST_INTERESTADUAL_KM else 'DIST_LONGE'

def _regiao_da_uf(uf):
    """Retorna a região de consolidação de uma UF."""
    for regiao, ufs in REGIOES_CONSOLIDACAO.items():
        if uf in ufs:
            return regiao
    return uf  # UF sem região → grupo próprio

# ── Classifica eng_final em SP vs interestadual ──────────────
print("\n🌎 Classificando destinos por UF...")

eng_sp             = []   # → pipeline diário normal
eng_interestadual  = []   # → plano interestadual

for _eng in eng_final:
    _uf = _uf_do_engradado(_eng)
    _eng['_uf_classificada'] = _uf
    if _uf in ('SP', '') or (_uf == 'DIST_LONGE' and
        _haversine((ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']),
                   (_eng.get('lat', 0), _eng.get('lon', 0)))[0] <= DIST_INTERESTADUAL_KM):
        eng_sp.append(_eng)
    else:
        eng_interestadual.append(_eng)

print(f"   📦 SP (rota diária)    : {len(eng_sp)} engradado(s)")
print(f"   ✈️  Interestadual       : {len(eng_interestadual)} engradado(s)")

# Substitui eng_final pelo subset SP para todo o pipeline diário abaixo
eng_final_original = eng_final   # preserva para referência
eng_final          = eng_sp      # pipeline usa só SP a partir daqui

# ── Plano Interestadual ───────────────────────────────────────
timeline_inter    = []
df_interestadual  = pd.DataFrame()

if eng_interestadual:
    print("\n   🗺️  Agrupando destinos interestaduais por região...")

    # Agrupa por região de consolidação
    _grupos_inter = {}
    for _eng in eng_interestadual:
        _uf  = _eng.get('_uf_classificada', 'OUT')
        _reg = _regiao_da_uf(_uf)
        _grupos_inter.setdefault(_reg, []).append(_eng)

    # Mostra resumo por região
    for _reg, _engs in sorted(_grupos_inter.items()):
        _ufs_reg = sorted(set(_e.get('_uf_classificada','?') for _e in _engs))
        print(f"   • {_reg:<16}: {len(_engs):>3} eng | UFs: {', '.join(_ufs_reg)}")

    # Tabela de frete já carregada — usa calcular_frete_viagem()
    _inter_trip_id = 1

    for _regiao, _engs_reg in sorted(_grupos_inter.items()):

        # Constrói grupos de destino para esta região
        _gdest_reg = {}
        for _e in _engs_reg:
            _cli = _e['cli']
            if _cli not in _gdest_reg:
                _gdest_reg[_cli] = {
                    'cli':     _cli,
                    'lat':     _e.get('lat', 0),
                    'lon':     _e.get('lon', 0),
                    'limite':  _e.get('limite', '23:59'),
                    'descarga': _e.get('descarga', TEMPO_DESCARGA_MIN),
                    'engs':    [],
                    'uf':      _e.get('_uf_classificada', '?'),
                }
            _gdest_reg[_cli]['engs'].append(_e)
        _dest_lista = list(_gdest_reg.values())

        # Tenta empacotar toda a região num único caminhão
        # Se não couber, subdivide por estado e depois por volume
        def _gerar_trips_inter(destinos):
            """Divide lista de destinos em trips respeitando capacidade."""
            trips = []
            pendentes = destinos[:]
            _iter = 0
            while pendentes and _iter < 50:
                _iter += 1
                # Tenta empacotar o máximo possível no melhor caminhão
                melhor_truck = None
                melhor_aloc  = []
                melhor_grupo = []
                # Começa com todos, vai reduzindo se não couber
                for _tam in range(len(pendentes), 0, -1):
                    _sub = pendentes[:_tam]
                    _trk, _aloc, _sob = empacotar_em_caminhao(_sub, truck_list)
                    if _trk and _aloc:
                        melhor_truck = _trk
                        melhor_aloc  = _aloc
                        melhor_grupo = _sub
                        break
                if not melhor_truck:
                    pendentes.pop(0)
                    continue
                trips.append((melhor_truck, melhor_grupo, melhor_aloc))
                clis_usadas = {g['cli'] for g in melhor_grupo}
                pendentes = [g for g in pendentes if g['cli'] not in clis_usadas]
            return trips

        # Ordena destinos: primeiro por UF, depois por limite de entrega
        _dest_lista.sort(key=lambda g: (g.get('uf',''), lim_min(g['limite'])))

        _trips_reg = _gerar_trips_inter(_dest_lista)

        print(f"\n   🚛 {_regiao}: {len(_trips_reg)} trip(s) interestadual(is)")

        _h_inter = datetime.today().replace(
            hour=HORA_INICIO, minute=0, second=0, microsecond=0)

        for _trk, _grp, _aloc in _trips_reg:
            _qtd_e  = sum(len(g['engs']) for g in _grp)
            _tc     = tempo_carregamento(_qtd_e)
            _hs     = _h_inter + timedelta(minutes=_tc)

            # Volume total alocado
            _vol_aloc = min(
                sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                    if 'dim_f' in e else e.get('vol', 0)
                    for e in _aloc),
                _trk['vol'])

            # Distância total estimada: Suzano → cada destino em ordem
            _km_total = 0.0
            _p_ref = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
            for _g in _grp:
                _km_g, _ = _haversine(_p_ref, (_g['lat'], _g['lon']))
                _km_total += _km_g
                _p_ref = (_g['lat'], _g['lon'])
            # Retorno ao armazém
            _km_ret, _ = _haversine(_p_ref,
                (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']))
            _km_total += _km_ret

            # Frete estimado (usa tabela de frete ou R$3,80/km fallback)
            try:
                _tf_inter = carregar_tabela_frete()
            except Exception:
                _tf_inter = {}
            _frete_est = (calcular_frete_viagem(_km_total, _trk['tipo'], _tf_inter)
                          if _tf_inter else _km_total * 3.80)

            _tid = f"INT_{_inter_trip_id:03d}"
            _ufs_trip = sorted(set(g.get('uf','?') for g in _grp))

            for _g in _grp:
                _engs_cli = [e for e in _aloc if e['cli'] == _g['cli']]
                _km_g, _ = _haversine(
                    (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']),
                    (_g['lat'], _g['lon']))

                timeline_inter.append({
                    'Viagem':              _tid,
                    'Região':              _regiao,
                    'UF':                  _g.get('uf', '?'),
                    'Caminhão':            _trk['tipo'],
                    'Vol Caminhão (m³)':   round(_trk['vol'], 2),
                    'Vol Carga Total (m³)':round(_vol_aloc, 4),
                    'Ocupação %':          round(_vol_aloc / _trk['vol'] * 100, 1),
                    'Cliente':             _g['cli'],
                    'Qtd Engradados':      len(_engs_cli),
                    'NFs':                 "/".join(sorted(set(
                                               e['nf'] for e in _engs_cli))),
                    'Distancia KM (est)':  round(_km_g, 1),
                    'Dist Total Trip (km)':round(_km_total, 1),
                    'Frete Estimado (R$)': round(_frete_est, 2),
                    'UFs na Viagem':       ", ".join(_ufs_trip),
                    'Limite Entrega':      _g['limite'],
                    'Observação':          'Transportadora terceirizada CIF/FOB',
                })

            _inter_trip_id += 1
            _h_inter = _hs  # próxima trip começa após carregamento desta

        # Sumário da região no console
        _nfs_reg = sorted(set(
            e['nf'] for e in _engs_reg if 'nf' in e))
        print(f"   📋 NFs: {', '.join(_nfs_reg[:8])}"
              f"{'...' if len(_nfs_reg) > 8 else ''}")
        _vol_reg = sum(e.get('vol', 0) for e in _engs_reg)
        print(f"   📦 Volume total: {_vol_reg:.3f} m³")

    # ── Salva Excel interestadual ─────────────────────────────
    df_interestadual = pd.DataFrame(timeline_inter)
    if not df_interestadual.empty:
        _path_inter = os.path.join(DRIVE_PATH, '06_TIMELINE_INTERESTADUAL.xlsx')
        try:
            with pd.ExcelWriter(_path_inter, engine='openpyxl') as _writer:
                # Aba geral
                df_interestadual.to_excel(_writer, sheet_name='Todas', index=False)
                # Uma aba por região
                for _reg in df_interestadual['Região'].unique():
                    _df_reg = df_interestadual[df_interestadual['Região'] == _reg]
                    _sheet  = _reg[:31]  # Excel limita 31 chars
                    _df_reg.to_excel(_writer, sheet_name=_sheet, index=False)
            print(f"\n   ✅ Plano interestadual → 06_TIMELINE_INTERESTADUAL.xlsx")
            print(f"   📊 {df_interestadual['Viagem'].nunique()} trip(s) | "
                  f"{len(df_interestadual)} destino(s) | "
                  f"Frete est.: R$ {df_interestadual['Frete Estimado (R$)'].sum():,.2f}")
        except Exception as _e_inter_xls:
            print(f"   ⚠️  Erro ao salvar interestadual: {repr(_e_inter_xls)}")
else:
    print("   ✅ Nenhum destino interestadual encontrado.")

# =============================================================
# 12. NÍVEL 2 — PLANEJAMENTO DE VIAGENS
# =============================================================
print("\n🚚 Nível 2: Planejando Viagens...")

# Passo A — Agrupa engradados por destino (resolve múltiplas NFs → 1 destino)
grupos = construir_grupos_destino(eng_final)
print(f"   📦 Engradados     : {len(eng_final)}")
print(f"   👥 Destinos únicos: {len(grupos)}")

# Passo B — Consolida destinos próximos em clusters de rota
clusters = consolidar_por_proximidade(grupos)
print(f"   🗺️  Clusters       : {len(clusters)} "
      f"(raio {RAIO_CONSOLIDACAO_KM} km)")

# Ordena clusters pela menor folga estimada:
# folga = limite_mais_restritivo - distancia_media_estimada
# Quem tem menos tempo sobrando sai primeiro
def _folga_cluster(c):
    lim   = min(lim_min(g['limite']) for g in c)
    # Estimativa rápida: distância haversine do armazém até o centróide do cluster
    lat_c = sum(g['lat'] for g in c) / len(c)
    lon_c = sum(g['lon'] for g in c) / len(c)
    _, mins_est = _haversine(
        (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']), (lat_c, lon_c)
    )
    return lim - mins_est   # menor folga = mais urgente

clusters = sorted(clusters, key=_folga_cluster)

timeline                = []
engradados_residuais    = []
id_trip                 = 1
# Duas filas de carregamento paralelas: Doca e Pátio
_t_inicio = datetime.today().replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
proxima_carga_doca  = _t_inicio   # Carreta, Truck, Toco
proxima_carga_patio = _t_inicio   # 608, Van, Fiorino

def _prox_carga(tipo_truck):
    """Retorna e atualiza a fila correta baseado no local do caminhão."""
    truck_info = next((t for t in truck_list if t['tipo'] == tipo_truck), None)
    local = truck_info['local'] if truck_info else 'Doca'
    return local

for cluster in clusters:
    # Ordena paradas do cluster: mais urgente primeiro
    paradas_ord = sorted(cluster, key=lambda g: lim_min(g['limite']))

    # Passo C — Verifica viabilidade temporal do cluster
    # Determina local de carregamento pelo menor caminhão que comporta o cluster
    _truck_est, _, _ = empacotar_em_caminhao(paradas_ord, truck_list)
    _local_cluster   = _truck_est['local'] if _truck_est else 'Doca'
    qtd_engs_cluster = sum(len(g['engs']) for g in paradas_ord)
    t_carga_cluster  = tempo_carregamento(qtd_engs_cluster)
    h_inicio_carga_c = proxima_carga_doca if _local_cluster == 'Doca' else proxima_carga_patio
    h_saida_cluster  = h_inicio_carga_c + timedelta(minutes=t_carga_cluster)
    rotas, viavel, chegadas = simular_viagem(
        paradas_ord, mem_cache, gmaps, API_DISPONIVEL, controlador_api,
        h_inicio=h_saida_cluster
    )

    if not viavel:
        # Cluster inviável — divide de forma inteligente usando APENAS o cache
        # (sem novas chamadas à API para não causar lentidão)
        # Estratégia: remove o último destino da rota sequencialmente
        # até o grupo restante ser viável. Usa chegadas já calculadas.
        def dividir_por_chegadas(paradas, chegadas_calc, h_saida):
            """
            Divide usando chegadas já calculadas — sem chamar API.
            Remove do final da rota os destinos que ultrapassam o limite.
            """
            viaveis   = []
            removidos = []
            for g in paradas:
                chegada = chegadas_calc.get(g['cli'])
                if chegada and dt_min(chegada) <= lim_min(g['limite']):
                    viaveis.append(g)
                else:
                    removidos.append(g)
            if not viaveis:
                return [[p] for p in paradas], []
            return [viaveis], removidos

        grupos_viaveis, removidos = dividir_por_chegadas(
            paradas_ord, chegadas, h_saida_cluster
        )
        n_rem = len(removidos)
        print(f"   ⚠️  Cluster {len(paradas_ord)} destinos inviável por janela "
              f"— {len(grupos_viaveis[0])} no prazo"
              f"{f', {n_rem} fora do prazo → trip(s) individual(is)' if n_rem else ''}.")
        # Removidos voltam como viagens individuais
        paradas_separadas = grupos_viaveis + [[r] for r in removidos]
    else:
        paradas_separadas = [paradas_ord]

    for grupo_viagem in paradas_separadas:
        # Roda simulação apenas para grupos individuais (grupos já viáveis
        # têm rotas e chegadas calculadas no simular_viagem acima)
        if len(grupo_viagem) == 1 and grupo_viagem[0] not in paradas_ord[:len(paradas_ord)]:
            qtd_e_gv = sum(len(g['engs']) for g in grupo_viagem)
            t_c_gv   = tempo_carregamento(qtd_e_gv)
            _local_gv = _prox_carga(empacotar_em_caminhao(grupo_viagem, truck_list)[0]['tipo'] if empacotar_em_caminhao(grupo_viagem, truck_list)[0] else 'Doca')
            h_s_gv    = (proxima_carga_doca if _local_gv == 'Doca' else proxima_carga_patio) + timedelta(minutes=t_c_gv)
            rotas, _, chegadas = simular_viagem(
                grupo_viagem, mem_cache, gmaps, API_DISPONIVEL, controlador_api,
                h_inicio=h_s_gv
            )

        # Passo D — Empacota engradados do grupo no menor caminhão possível
        truck, alocados, sobram = empacotar_em_caminhao(grupo_viagem, truck_list)

        if not alocados:
            print(f"   ❌ Sem caminhão para cluster — "
                  f"{sum(len(g['engs']) for g in grupo_viagem)} eng. não alocados.")
            engradados_residuais.extend(
                [eng for g in grupo_viagem for eng in g['engs']]
            )
            continue

        if sobram:
            print(f"   ⚠️  {len(sobram)} eng. excedem o caminhão → nova viagem.")
            # Sobras voltam como grupos individuais para reprocessamento
            for eng_sob in sobram:
                engradados_residuais.append(eng_sob)

        # Passo E — Monta linhas do timeline
        # Agrupa alocados por cliente para consolidar NFs
        cli_map = {}
        for eng in alocados:
            cli_map.setdefault(eng['cli'], []).append(eng)

        # Volume físico ocupado no caminhão:
        # usa dim_f (dimensão real após posicionamento pelo tetris)
        # cap em truck['vol'] para garantir que ocupação nunca ultrapasse 100%
        vol_truck      = truck['vol']
        vol_total_trip = min(
            sum(
                (e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                if 'dim_f' in e else (e['dim'][0]*e['dim'][1]*e['dim'][2])
                for e in alocados
            ),
            vol_truck
        )

        # Carregamento paralelo: Doca e Pátio independentes
        qtd_engs_trip  = sum(len(g['engs']) for g in grupo_viagem)
        t_carga        = tempo_carregamento(qtd_engs_trip)
        local_truck    = _prox_carga(truck['tipo'])
        if local_truck == 'Pateo':
            h_inicio_carga   = proxima_carga_patio
            h_ref_saida      = h_inicio_carga + timedelta(minutes=t_carga)
            proxima_carga_patio = h_ref_saida
        else:
            h_inicio_carga   = proxima_carga_doca
            h_ref_saida      = h_inicio_carga + timedelta(minutes=t_carga)
            proxima_carga_doca = h_ref_saida
        h_ref = h_ref_saida
        p_ref = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])

        for g in grupo_viagem:
            cli  = g['cli']
            engs = cli_map.get(cli, [])
            if not engs:
                continue

            dest          = (g['lat'], g['lon'])
            rid           = _rota_id(p_ref, dest)
            km, mins      = rotas.get(rid, _haversine(p_ref, dest))
            chegada_bruta = h_ref + timedelta(minutes=mins)
            chegada       = ajustar_pausas(h_ref, chegada_bruta)

            em_cache   = rid in mem_cache
            fonte_rota = ('Cache'      if em_cache else
                          'Google API' if API_DISPONIVEL and not controlador_api.bloqueado
                          else 'Haversine')

            timeline.append({
                'Viagem':              f"TRIP_{id_trip:03d}",
                'Caminhão':            truck['tipo'],
                'Vol Caminhão (m³)':   round(truck['vol'], 2),
                'Vol Carga Total (m³)': round(vol_total_trip, 4),
                'Ocupação %':          round(vol_total_trip / truck['vol'] * 100, 1),
                'Cliente':             cli,
                'Qtd Engradados':      len(engs),
                'Tipos Engradados':    ", ".join(sorted(set(e['tipo'] for e in engs))),
                'NFs':                 "/".join(sorted(set(e['nf'] for e in engs))),
                'Início Carga':        h_inicio_carga.strftime("%H:%M"),
                'Local Carga':         local_truck,
                'Tempo Carga (min)':   t_carga,
                'Saída Armazém':       h_ref_saida.strftime("%H:%M"),
                'Distancia KM':        round(km, 2),
                'Chegada':             chegada.strftime("%H:%M"),
                'Tempo Descarga (min)': g.get('descarga', TEMPO_DESCARGA_MIN),
                'Limite Entrega':      g['limite'],
                'Dentro do Prazo':     '✅' if dt_min(chegada) <= lim_min(g['limite'])
                                       else '❌',
                'Fonte Rota':          fonte_rota
            })

            # Após descarga: usa tempo individual do cliente
            _desc_g = g.get('descarga', TEMPO_DESCARGA_MIN)
            saida_proxima = chegada + timedelta(minutes=_desc_g)
            h_ref = ajustar_pausas(chegada, saida_proxima)
            p_ref = dest

        id_trip += 1

# Passo F — Processa engradados residuais (sobras de empacotamento)
if engradados_residuais:
    print(f"\n   ♻️  {len(engradados_residuais)} engradado(s) residuais → viagens avulsas.")
    grupos_res = construir_grupos_destino(engradados_residuais)
    for g in grupos_res:
        truck, alocados, _ = empacotar_em_caminhao([g], truck_list)
        if not alocados:
            print(f"   ❌ Eng. residual '{g['cli']}' sem caminhão — descartado.")
            continue

        dest    = (g['lat'], g['lon'])
        p_arm   = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
        rotas_r = obter_rotas_lote(
            [(p_arm, dest)], mem_cache, gmaps, API_DISPONIVEL, controlador_api
        )
        rid      = _rota_id(p_arm, dest)
        km, mins = rotas_r.get(rid, _haversine(p_arm, dest))
        chegada  = datetime.today().replace(
            hour=HORA_INICIO, minute=0, second=0, microsecond=0
        ) + timedelta(minutes=mins)

        vol_carga = min(
            sum(
                (e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                if 'dim_f' in e else (e['dim'][0]*e['dim'][1]*e['dim'][2])
                for e in alocados
            ),
            truck['vol']
        )
        timeline.append({
            'Viagem':              f"TRIP_{id_trip:03d}",
            'Caminhão':            truck['tipo'],
            'Vol Caminhão (m³)':   round(truck['vol'], 2),
            'Vol Carga Total (m³)': round(vol_carga, 4),
            'Ocupação %':          round(vol_carga / truck['vol'] * 100, 1),
            'Cliente':             g['cli'],
            'Qtd Engradados':      len(alocados),
            'Tipos Engradados':    ", ".join(sorted(set(e['tipo'] for e in alocados))),
            'NFs':                 "/".join(sorted(g['nfs'])),
            'Distancia KM':        round(km, 2),
            'Chegada':             chegada.strftime("%H:%M"),
            'Limite Entrega':      g['limite'],
            'Dentro do Prazo':     '✅' if dt_min(chegada) <= lim_min(g['limite'])
                                   else '❌',
            'Fonte Rota':          'Haversine'
        })
        id_trip += 1




# =============================================================
# VISUALIZAÇÃO 3D — ENGRADADO E CAMINHÃO
# =============================================================

def _cubo_wireframe(ax, x, y, z, dx, dy, dz, cor, alpha=0.75, lw=0.5):
    """Desenha um cubo sólido + wireframe de arestas no eixo 3D."""
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    verts = [
        [(x,y,z),     (x+dx,y,z),     (x+dx,y+dy,z),     (x,y+dy,z)    ],  # base
        [(x,y,z+dz),  (x+dx,y,z+dz),  (x+dx,y+dy,z+dz),  (x,y+dy,z+dz)],  # topo
        [(x,y,z),     (x+dx,y,z),     (x+dx,y,z+dz),     (x,y,z+dz)    ],  # frente
        [(x,y+dy,z),  (x+dx,y+dy,z),  (x+dx,y+dy,z+dz),  (x,y+dy,z+dz)],  # fundo
        [(x,y,z),     (x,y+dy,z),     (x,y+dy,z+dz),     (x,y,z+dz)    ],  # esq
        [(x+dx,y,z),  (x+dx,y+dy,z),  (x+dx,y+dy,z+dz),  (x+dx,y,z+dz)],  # dir
    ]
    col = Poly3DCollection(verts, alpha=alpha, linewidths=lw,
                           edgecolors='#333333', facecolors=cor)
    ax.add_collection3d(col)


def _paleta_pns(pns_unicos):
    """Gera dict {pn: cor_hex} com cores distintas para cada part number."""
    import colorsys
    n     = max(len(pns_unicos), 1)
    cores = {}
    for i, pn in enumerate(sorted(pns_unicos)):
        h = i / n
        r, g, b = colorsys.hsv_to_rgb(h, 0.72, 0.88)
        cores[pn] = (r, g, b)
    return cores


def gerar_imagem_engradado(eng_id, lay, escolha, caminho_saida):
    """
    Gera visualização 3D do engradado com:
      - Uma cor por part number
      - Legenda com PN e cor
      - Título com ID sequência, tipo e eficiência
      - Contorno translúcido do engradado (container)
    """
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D   # noqa: F401

    pns_unicos = sorted(set(p['pn'] for p in lay))
    cores      = _paleta_pns(pns_unicos)

    fig = plt.figure(figsize=(10, 7))
    ax  = fig.add_subplot(111, projection='3d')

    C, L, A = escolha['dim']

    # Contorno do engradado (wireframe transparente)
    _cubo_wireframe(ax, 0, 0, 0, C, L, A,
                    cor=(0.85, 0.85, 0.85), alpha=0.08, lw=1.0)

    # Peças
    for p in lay:
        if 'pos' not in p or 'dim_f' not in p:
            continue
        px, py, pz     = p['pos']
        dc, dl, da     = p['dim_f']
        cor            = cores[p['pn']]
        _cubo_wireframe(ax, px, py, pz, dc, dl, da, cor=cor, alpha=0.82)

    # Eixos e limites
    ax.set_xlim(0, C); ax.set_ylim(0, L); ax.set_zlim(0, A)
    ax.set_xlabel('Comprimento (m)', labelpad=6)
    ax.set_ylabel('Largura (m)',     labelpad=6)
    ax.set_zlabel('Altura (m)',      labelpad=6)
    ax.view_init(elev=22, azim=-55)

    # Legenda de part numbers
    from matplotlib.patches import Patch
    handles = [Patch(facecolor=cores[pn], edgecolor='#444', label=pn)
               for pn in pns_unicos]
    ax.legend(handles=handles, loc='upper left', fontsize=7,
              title='Part Numbers', title_fontsize=8,
              bbox_to_anchor=(0.0, 1.0))

    # Volume e eficiência
    vol_pecas  = min(
        sum(min(p['dim_f'][0]*p['dim_f'][1]*p['dim_f'][2], escolha['vol'])
            for p in lay if 'dim_f' in p),
        escolha['vol']
    )
    eficiencia = round(vol_pecas / escolha['vol'] * 100, 1)

    titulo_eng = (
        f"Engradado {eng_id}  |  {escolha['nome']}"
        f"  ({C:.2f} x {L:.2f} x {A:.2f} m)\n"
        f"{len(lay)} pecas  |  {len(pns_unicos)} PNs"
        f"  |  Eficiencia volumetrica: {eficiencia}%"
    )
    ax.set_title(titulo_eng, fontsize=10, pad=12)

    plt.tight_layout()
    plt.savefig(caminho_saida, dpi=150, bbox_inches='tight')
    plt.show()
    plt.close(fig)
    print(f"   🖼️  Engradado 3D salvo → {caminho_saida}")


def _cilindro_3d(ax, cx, cy, cz, raio, largura, cor, alpha=0.9, n=20):
    """Desenha um cilindro (roda) deitado no eixo Y."""
    import numpy as np
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    theta  = np.linspace(0, 2*np.pi, n)
    ys     = [cy, cy + largura]
    verts  = []
    # Faces laterais
    for i in range(n - 1):
        t0, t1 = theta[i], theta[i+1]
        for y0, y1 in [(ys[0], ys[1])]:
            quad = [
                (cx + raio*np.cos(t0), y0, cz + raio*np.sin(t0)),
                (cx + raio*np.cos(t1), y0, cz + raio*np.sin(t1)),
                (cx + raio*np.cos(t1), y1, cz + raio*np.sin(t1)),
                (cx + raio*np.cos(t0), y1, cz + raio*np.sin(t0)),
            ]
            verts.append(quad)
    # Tampas
    for y in ys:
        tampa = [(cx + raio*np.cos(t), y, cz + raio*np.sin(t)) for t in theta]
        verts.append(tampa)
    col = Poly3DCollection(verts, alpha=alpha, linewidths=0.2,
                           edgecolors='#222', facecolors=cor)
    ax.add_collection3d(col)


def gerar_imagem_caminhao(trip_id, truck, alocados_caminhao, caminho_saida):
    """
    Gera visualização 3D do baú do caminhão com engradados coloridos por cliente.
    Cabine e rodas desativadas — exibe apenas a área útil de carga.
    """
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D          # noqa: F401
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    from matplotlib.patches import Patch

    clientes_unicos = sorted(set(e['cli'] for e in alocados_caminhao))
    cores           = _paleta_pns(clientes_unicos)

    C, L, A = truck['dim']

    fig = plt.figure(figsize=(13, 7))
    ax  = fig.add_subplot(111, projection='3d')
    ax.set_facecolor('#f4f4f4')
    fig.patch.set_facecolor('#f4f4f4')

    # Contorno do baú (área útil de carga)
    _cubo_wireframe(ax, 0, 0, 0, C, L, A,
                    cor=(0.55, 0.75, 0.95), alpha=0.07, lw=1.4)

    # Piso do baú
    piso = [[(0,0,0),(C,0,0),(C,L,0),(0,L,0)]]
    ax.add_collection3d(Poly3DCollection(piso,
                        alpha=0.25, facecolors='#ccc',
                        edgecolors='#888', linewidths=0.5))

    # Engradados
    for eng in alocados_caminhao:
        if 'pos' not in eng or 'dim_f' not in eng:
            continue
        px, py, pz = eng['pos']
        dc, dl, da = eng['dim_f']
        _cubo_wireframe(ax, px, py, pz, dc, dl, da,
                        cor=cores[eng['cli']], alpha=0.88)

    ax.set_xlim(0, C)
    ax.set_ylim(0, L)
    ax.set_zlim(0, A + 0.1)
    ax.set_xlabel('Comprimento (m)', labelpad=6, fontsize=8)
    ax.set_ylabel('Largura (m)',     labelpad=6, fontsize=8)
    ax.set_zlabel('Altura (m)',      labelpad=6, fontsize=8)
    ax.view_init(elev=20, azim=-50)

    ax.xaxis.pane.fill = False
    ax.yaxis.pane.fill = False
    ax.zaxis.pane.fill = False
    ax.xaxis.pane.set_edgecolor('#cccccc')
    ax.yaxis.pane.set_edgecolor('#cccccc')
    ax.zaxis.pane.set_edgecolor('#cccccc')
    ax.grid(True, alpha=0.12, linewidth=0.4)

    # ── Legenda ─────────────────────────────────────────────
    handles = [Patch(facecolor=cores[c], edgecolor='#444',
                     label=c.split('[')[0].strip()[:32] +
                           ('…' if len(c.split('[')[0].strip()) > 32 else ''))
               for c in clientes_unicos]
    ax.legend(handles=handles, loc='upper left', fontsize=7,
              title='Clientes', title_fontsize=8,
              bbox_to_anchor=(0.0, 1.0), framealpha=0.88)

    # ── Título ──────────────────────────────────────────────
    vol_carga = min(
        sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
            if 'dim_f' in e else e['vol']
            for e in alocados_caminhao),
        truck['vol']
    )
    ocupacao = round(vol_carga / truck['vol'] * 100, 1)
    n_engs   = len(alocados_caminhao)
    n_half   = sum(1 for e in alocados_caminhao if e.get('tipo') == 'Half')
    n_full   = n_engs - n_half
    ax.set_title(
        f"Viagem {trip_id}  |  {truck['tipo']}"
        f"  ({C:.1f} × {L:.1f} × {A:.1f} m baú)\n"
        f"{n_engs} engradados ({n_half} Half / {n_full} Full)"
        f"  |  Ocupação: {ocupacao}%",
        fontsize=10, pad=14
    )

    plt.tight_layout()
    plt.savefig(caminho_saida, dpi=150, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    plt.show()
    plt.close(fig)
    print(f"   🖼️  Caminhão 3D salvo → {caminho_saida}")

# =============================================================
# FUNÇÃO UTILITÁRIA — salvar Excel com retry (evita PermissionError)
# =============================================================
def salvar_excel(df, caminho, descricao='arquivo'):
    """Salva DataFrame em Excel. Se o arquivo estiver aberto, aguarda e tenta novamente."""
    import time as _time
    for tentativa in range(1, 4):
        try:
            df.to_excel(caminho, index=False)
            return True
        except PermissionError:
            if tentativa < 3:
                print(f"   ⚠️  {descricao} está aberto — feche o Excel e aguarde 5s "
                      f"(tentativa {tentativa}/3)...")
                _time.sleep(5)
            else:
                print(f"   ❌ Não foi possível salvar {descricao}. "
                      f"Feche o arquivo no Excel e rode novamente.")
                return False
    return False

# =============================================================
# CARREGA CONSUMO DE COMBUSTÍVEL
# =============================================================
def carregar_consumo_combustivel():
    """
    Lê Consumo Combustivel.xlsx:
      Linha 1: tipos de caminhão
      Linha 2: consumo em km/litro (ou litros/100km — detecta automaticamente)
    Retorna dict {tipo_caminhao: litros_por_km}
    """
    caminho_cc = os.path.join(DRIVE_PATH, 'Consumo Combustivel.xlsx')
    consumo = {}
    try:
        # dtype=str garante que '608' não seja lido como inteiro no header
        df_cc = pd.read_excel(caminho_cc, header=None, dtype=str)
        tipos   = df_cc.iloc[0].astype(str).str.strip().tolist()
        valores = df_cc.iloc[1].tolist()
        for tipo, val in zip(tipos, valores):
            tipo_str = str(tipo).strip()
            if not tipo_str or tipo_str.lower() in ('nan', 'none', ''):
                continue
            try:
                v = float(val)
                if v >= 1:
                    litros_por_km = 1.0 / v   # km/L → L/km
                else:
                    litros_por_km = v          # já em L/km
                consumo[tipo_str] = litros_por_km
            except (ValueError, ZeroDivisionError):
                pass
        print(f"   ⛽ Consumo carregado: {len(consumo)} caminhões")
    except FileNotFoundError:
        print(f"   ⚠️  Consumo Combustivel.xlsx não encontrado em {DRIVE_PATH}")
    except Exception as e:
        print(f"   ⚠️  Erro ao ler Consumo Combustivel.xlsx: {repr(e)}")
    return consumo

# Fator de emissão diesel: 2.676 kg CO2 por litro (IPCC/EPA)
CO2_KG_POR_LITRO = 2.676

def calcular_co2(df_timeline, consumo_combustivel):
    """
    Calcula emissão de CO2 por VIAGEM (não por parada).
    A distância de uma viagem é a distância total percorrida —
    soma das distâncias de cada perna (armazém→p1→p2→...→pN).
    O CO2 é calculado uma vez por viagem e distribuído igualmente
    entre as paradas para fins de relatório.
    Retorna DataFrame com coluna CO2 (kg) e total em kg e toneladas.
    """
    if df_timeline.empty or not consumo_combustivel:
        return df_timeline, 0.0, 0.0

    df = df_timeline.copy()

    # Agrupa distâncias por viagem (soma das pernas)
    dist_por_viagem = df.groupby('Viagem')['Distancia KM'].sum().to_dict()

    # Identifica o caminhão de cada viagem
    caminhao_col = 'Caminhão' if 'Caminhão' in df.columns else 'Caminhao'
    truck_por_viagem = df.groupby('Viagem')[caminhao_col].first().to_dict()

    # CO2 por viagem
    co2_por_viagem = {}
    for trip, km_total in dist_por_viagem.items():
        tipo = str(truck_por_viagem.get(trip, '')).strip()   # garante string
        lkm  = consumo_combustivel.get(tipo)
        if lkm is None:
            for chave in consumo_combustivel:
                chave_s = str(chave).strip()
                tipo_s  = str(tipo).strip()
                if tipo_s and chave_s and (
                        tipo_s.lower() in chave_s.lower() or
                        chave_s.lower() in tipo_s.lower()):
                    lkm = consumo_combustivel[chave]
                    break
        co2_por_viagem[trip] = round(km_total * (lkm or 0) * CO2_KG_POR_LITRO, 2)

    # Distribui CO2 igualmente entre paradas da mesma viagem
    n_paradas_por_viagem = df.groupby('Viagem').size().to_dict()
    def co2_linha(row):
        trip  = row.get('Viagem', '')
        n     = n_paradas_por_viagem.get(trip, 1)
        return round(co2_por_viagem.get(trip, 0) / n, 2)

    df['CO2 (kg)'] = df.apply(co2_linha, axis=1)
    total_kg = sum(co2_por_viagem.values())
    return df, total_kg, total_kg / 1000

# =============================================================
# 13. SAÍDAS FINAIS
# =============================================================
consumo_combustivel = carregar_consumo_combustivel()

df_tl = pd.DataFrame(timeline)

# Adiciona CO2 ao timeline real
df_tl, co2_real_kg, co2_real_t = calcular_co2(df_tl, consumo_combustivel)

salvar_excel(df_tl, os.path.join(DRIVE_PATH, '06_TIMELINE_REAL.xlsx'), '06_TIMELINE_REAL.xlsx')
salvar_cache(CACHE_FILE, mem_cache)

# Cálculos do timeline real — feitos ANTES do otimizado para uso no comparativo
viagens          = id_trip - 1
no_prazo         = int((df_tl['Dentro do Prazo'] == '✅').sum()) if not df_tl.empty else 0
total_linhas     = len(df_tl)
n_caminhoes_real = df_tl['Viagem'].nunique() if not df_tl.empty else 0
cli_por_viagem_real = df_tl.groupby('Viagem')['Cliente'].count().mean() if not df_tl.empty else 0

print(f"\n   ✅ {viagens} viagem(ns) → 06_TIMELINE_REAL.xlsx")
if not df_tl.empty:
    print(f"   📋 NFs consolidadas   : "
          f"{df_tl['NFs'].str.split('/').apply(len).sum()} NFs em {viagens} viagens")
    print(f"   ⏱️  Entregas no prazo  : {no_prazo}/{total_linhas} "
          f"({round(no_prazo/total_linhas*100,1) if total_linhas else 0}%)")
    print(f"   🚛 Caminhões usados    : {n_caminhoes_real}")
    print(f"   🚛 Ocupação média      : {df_tl['Ocupação %'].mean():.1f}%")
    if co2_real_kg > 0:
        print(f"   🌿 Emissão CO2 est.   : {co2_real_kg:.1f} kg ({co2_real_t:.3f} t)")

# =============================================================
# TIMELINE OTIMIZADO — limite 02:00 do dia seguinte para todos
# =============================================================
print("\n🔄 Gerando Timeline Otimizado (limite 02:00 do dia seguinte)...")

# Sobrescreve limite de todos os grupos para '02:00+1' (1560 min)
LIMITE_OTIMIZADO = '02:00+1'

grupos_otm  = construir_grupos_destino(eng_final)
for g in grupos_otm:
    # Usa o limite negociado do cliente; fallback para LIMITE_OTIMIZADO
    g['limite'] = limite_otimizado_cliente(g['cli'])

clusters_otm = consolidar_por_proximidade(grupos_otm)

# Ordena clusters otimizados pela menor folga (limite negociado - dist estimada)
def _folga_cluster_otm(c):
    lim   = min(lim_min(g['limite']) for g in c)
    lat_c = sum(g['lat'] for g in c) / len(c)
    lon_c = sum(g['lon'] for g in c) / len(c)
    _, mins_est = _haversine(
        (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']), (lat_c, lon_c)
    )
    return lim - mins_est

clusters_otm = sorted(clusters_otm, key=_folga_cluster_otm)

timeline_otm  = []
id_trip_otm   = 1
eng_res_otm   = []
# Duas filas paralelas no otimizado
_t_ini_otm          = datetime.today().replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
proxima_carga_doca_otm  = _t_ini_otm
proxima_carga_patio_otm = _t_ini_otm
proxima_carga_otm       = _t_ini_otm   # alias mantido para subdividir_cluster

def subdividir_cluster(paradas, h_inicio, cache, gmaps_c, api_ok, ctrl):
    """
    Divide um cluster em subgrupos viáveis para o limite 02:00+1.
    Estratégia binária: testa o cluster inteiro; se inviável, divide ao meio
    e testa cada metade recursivamente. Garante que cada subgrupo chegue
    antes de 02:00 do dia seguinte.
    """
    if not paradas:
        return []
    # Testa o grupo completo
    try:
        _, viavel, _ = simular_viagem(paradas, cache, gmaps_c, api_ok, ctrl,
                                       h_inicio=h_inicio)
    except Exception:
        viavel = False

    if viavel or len(paradas) == 1:
        return [paradas]   # cabe tudo numa viagem

    # Não cabe — divide ao meio e testa cada parte
    meio   = len(paradas) // 2
    parte1 = subdividir_cluster(paradas[:meio], h_inicio, cache, gmaps_c, api_ok, ctrl)
    parte2 = subdividir_cluster(paradas[meio:], h_inicio, cache, gmaps_c, api_ok, ctrl)
    return parte1 + parte2


for idx_cluster, cluster in enumerate(clusters_otm):
    paradas_ord = sorted(cluster, key=lambda g: lim_min(g['limite']))

    # Calcula tempo de carga deste cluster para estimar saída
    qtd_engs_cl  = sum(len(g['engs']) for g in paradas_ord)
    t_carga_cl   = tempo_carregamento(qtd_engs_cl)
    h_saida_cl   = proxima_carga_otm + timedelta(minutes=t_carga_cl)

    # Subdivide o cluster em subgrupos que respeitam o limite negociado
    subgrupos = subdividir_cluster(
        paradas_ord, h_saida_cl, mem_cache, gmaps, API_DISPONIVEL, controlador_api
    )
    if len(subgrupos) > 1:
        print(f"   ℹ️  Cluster {idx_cluster+1}: dividido em {len(subgrupos)} viagens "
              f"para respeitar limite negociado.")

    for grupo_viagem in subgrupos:
        # Empacota primeiro para saber o caminhão e o local de carregamento
        try:
            truck_o, alocados_o, sobram_o = empacotar_em_caminhao(grupo_viagem, truck_list)
        except Exception as e_emp:
            print(f"   ⚠️  Erro empacotar otimizado: {repr(e_emp)}")
            eng_res_otm.extend([eng for g in grupo_viagem for eng in g['engs']])
            continue

        if not alocados_o:
            eng_res_otm.extend([eng for g in grupo_viagem for eng in g['engs']])
            continue

        # Carregamento paralelo Doca/Pátio — usa fila correta pelo tipo do caminhão
        qtd_engs_sg = sum(len(g['engs']) for g in grupo_viagem)
        t_carga_sg  = tempo_carregamento(qtd_engs_sg)
        local_sg    = _prox_carga(truck_o['tipo'])
        if local_sg == 'Pateo':
            h_otm_inicio            = proxima_carga_patio_otm
            h_saida_sg              = h_otm_inicio + timedelta(minutes=t_carga_sg)
            proxima_carga_patio_otm = h_saida_sg
        else:
            h_otm_inicio           = proxima_carga_doca_otm
            h_saida_sg             = h_otm_inicio + timedelta(minutes=t_carga_sg)
            proxima_carga_doca_otm = h_saida_sg
        proxima_carga_otm = min(proxima_carga_doca_otm, proxima_carga_patio_otm)

        # Recalcula rotas a partir da saída real
        try:
            rotas_o, _, chegadas_o = simular_viagem(
                grupo_viagem, mem_cache, gmaps, API_DISPONIVEL, controlador_api,
                h_inicio=h_saida_sg
            )
        except Exception as e_sim:
            print(f"   ⚠️  Erro simular_viagem otimizado: {repr(e_sim)}")
            eng_res_otm.extend([eng for g in grupo_viagem for eng in g['engs']])
            continue

        if sobram_o:
            eng_res_otm.extend(sobram_o)

        if not alocados_o:
            eng_res_otm.extend([eng for g in grupo_viagem for eng in g['engs']])
            continue

        if sobram_o:
            eng_res_otm.extend(sobram_o)

        cli_map_o = {}
        for eng in alocados_o:
            cli_map_o.setdefault(eng['cli'], []).append(eng)

        vol_total_o = min(
            sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                if 'dim_f' in e else (e['dim'][0]*e['dim'][1]*e['dim'][2])
                for e in alocados_o),
            truck_o['vol']
        )

        h_ref = h_saida_sg
        p_ref = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])

        for g in grupo_viagem:
            cli  = g['cli']
            engs = cli_map_o.get(cli, [])
            if not engs:
                continue
            dest          = (g['lat'], g['lon'])
            rid           = _rota_id(p_ref, dest)
            km, mins      = rotas_o.get(rid, _haversine(p_ref, dest))
            chegada_bruta = h_ref + timedelta(minutes=mins)
            chegada       = ajustar_pausas(h_ref, chegada_bruta)
            no_prazo      = dt_min(chegada) <= lim_min(g['limite'])

            timeline_otm.append({
                'Viagem':               f"TRIP_{id_trip_otm:03d}",
                'Caminhão':             truck_o['tipo'],
                'Vol Caminhão (m³)':    round(truck_o['vol'], 2),
                'Vol Carga Total (m³)': round(vol_total_o, 4),
                'Ocupação %':           round(vol_total_o / truck_o['vol'] * 100, 1),
                'Cliente':              cli,
                'Qtd Engradados':       len(engs),
                'Tipos Engradados':     ", ".join(sorted(set(e['tipo'] for e in engs))),
                'NFs':                  "/".join(sorted(set(e['nf'] for e in engs))),
                'Início Carga':         h_otm_inicio.strftime("%H:%M"),
                'Local Carga':          local_sg,
                'Tempo Carga (min)':    t_carga_sg,
                'Saída Armazém':        h_saida_sg.strftime("%H:%M"),
                'Distancia KM':         round(km, 2),
                'Chegada':              chegada.strftime("%H:%M"),
                'Dia Entrega':          'Mesmo dia' if chegada.date() == datetime.today().date()
                                        else 'Dia seguinte',
                'Limite Entrega':       g['limite'],
                'Dentro do Prazo':      '✅' if no_prazo else '❌',
                'Fonte Rota':           'Cache' if rid in mem_cache else
                                        ('Google API' if API_DISPONIVEL else 'Haversine')
            })
            _desc_otm = g.get('descarga', TEMPO_DESCARGA_MIN)
            saida_proxima_otm = chegada + timedelta(minutes=_desc_otm)
            h_ref = ajustar_pausas(chegada, saida_proxima_otm)
            p_ref = dest

        id_trip_otm += 1

# Residuais do otimizado
if eng_res_otm:
    grupos_res_otm = construir_grupos_destino(eng_res_otm)
    for g in grupos_res_otm:
        g['limite'] = limite_otimizado_cliente(g['cli'])
        truck_o, alocados_o, _ = empacotar_em_caminhao([g], truck_list)
        if not alocados_o:
            continue
        dest    = (g['lat'], g['lon'])
        p_arm   = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
        rotas_r = obter_rotas_lote([(p_arm, dest)], mem_cache, gmaps,
                                   API_DISPONIVEL, controlador_api)
        rid      = _rota_id(p_arm, dest)
        km, mins = rotas_r.get(rid, _haversine(p_arm, dest))
        chegada  = datetime.today().replace(
            hour=HORA_INICIO, minute=0, second=0, microsecond=0
        ) + timedelta(minutes=mins)
        vol_o = min(
            sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                if 'dim_f' in e else e['vol'] for e in alocados_o),
            truck_o['vol']
        )
        timeline_otm.append({
            'Viagem':               f"TRIP_{id_trip_otm:03d}",
            'Caminhão':             truck_o['tipo'],
            'Vol Caminhão (m³)':    round(truck_o['vol'], 2),
            'Vol Carga Total (m³)': round(vol_o, 4),
            'Ocupação %':           round(vol_o / truck_o['vol'] * 100, 1),
            'Cliente':              g['cli'],
            'Qtd Engradados':       len(alocados_o),
            'Tipos Engradados':     ", ".join(sorted(set(e['tipo'] for e in alocados_o))),
            'NFs':                  "/".join(sorted(g['nfs'])),
            'Distancia KM':         round(km, 2),
            'Chegada':              chegada.strftime("%H:%M"),
            'Dia Entrega':          'Mesmo dia' if chegada.date() == datetime.today().date()
                                    else 'Dia seguinte',
            'Limite Entrega':       g['limite'],
            'Dentro do Prazo':      '✅',
            'Fonte Rota':           'Haversine'
        })
        id_trip_otm += 1

df_otm = pd.DataFrame(timeline_otm)
df_otm, co2_otm_kg, co2_otm_t = calcular_co2(df_otm, consumo_combustivel)
salvar_excel(df_otm, os.path.join(DRIVE_PATH, '06_TIMELINE_OTIMIZADO.xlsx'),
             '06_TIMELINE_OTIMIZADO.xlsx')

# =============================================================
# SEÇÃO 13C — ALGORITMO GENÉTICO (AG)
# Otimiza o plano de rotas evoluindo uma população de soluções.
#
# Cromossomo : permutação dos índices dos grupos de destino
#              (cada gene = um destino; a ordem define como
#               eles são agrupados em trips)
# Fitness    : minimiza frete total + penalidade por atraso
# Operadores : crossover OX (Order Crossover) + mutação 2-opt
# =============================================================

# ── Parâmetros do AG (ajustáveis) ──────────────────────────
AG_POPULACAO      = 120    # indivíduos por geração
AG_GERACOES       = 80     # número de gerações
AG_TAXA_CROSSOVER = 0.85   # probabilidade de crossover
AG_TAXA_MUTACAO   = 0.15   # probabilidade de mutação por indivíduo
AG_ELITISMO       = 6      # melhores indivíduos preservados sem alteração
AG_PENALIDADE_ATR = 500.0  # R$ por entrega fora do prazo
AG_MAX_PARADAS    = 8      # máximo de destinos por trip (controla tamanho do cluster)
AG_USAR_CACHE_API = True   # usa cache existente (sem novas chamadas à API)



def carregar_tabela_frete():
    caminho = os.path.join(DRIVE_PATH, 'Tabela de Frete.xlsx')
    tabela  = {}
    try:
        # dtype=str garante que '608' não seja lido como inteiro
        df_fr      = pd.read_excel(caminho, header=0, dtype=str)
        # Converte valores numéricos de volta para float onde necessário
        col_faixa  = df_fr.columns[0]
        cols_trucks = df_fr.columns[1:]
        for col in cols_trucks:
            tipo   = str(col).strip()
            faixas = []
            for _, row in df_fr.iterrows():
                faixa = str(row[col_faixa]).strip()
                try:
                    valor = float(row[col])
                except (ValueError, TypeError):
                    continue
                if faixa.startswith(">"):
                    dist_max = 999999
                else:
                    partes   = re.split(r"[-]", faixa)
                    dist_max = float(partes[-1]) if partes else 999999
                faixas.append((dist_max, valor))
            if faixas:
                tabela[tipo] = sorted(faixas, key=lambda x: x[0])
        print(f"   Tabela de frete: {len(tabela)} caminhoes x {len(df_fr)} faixas")
    except FileNotFoundError:
        print(f"   Tabela de Frete.xlsx nao encontrada em {DRIVE_PATH}")
    except Exception as e:
        print(f"   Erro ao ler Tabela de Frete.xlsx: {repr(e)}")
    return tabela

def calcular_frete_viagem(km, tipo_caminhao, tabela_frete):
    """
    Frete = km × valor_por_km da faixa correspondente.
    A tabela tem o valor unitário (R$/km) para cada faixa de distância.
    Exemplo: 34 km com VAN a R$16/km = R$544.
    """
    tipo_caminhao = str(tipo_caminhao).strip()
    faixas = tabela_frete.get(tipo_caminhao)
    if not faixas:
        for chave in tabela_frete:
            chave_s = str(chave).strip()
            if (tipo_caminhao.lower() in chave_s.lower() or
                    chave_s.lower() in tipo_caminhao.lower()):
                faixas = tabela_frete[chave]
                break
    if not faixas:
        return 0.0
    # Encontra a faixa correta e multiplica km × valor_por_km
    for dist_max, valor_por_km in faixas:
        if km <= dist_max:
            return round(km * valor_por_km, 2)
    return round(km * faixas[-1][1], 2)   # última faixa para distâncias acima do máximo

def calcular_frete_total(df_timeline, tabela_frete):
    """
    Calcula frete por VIAGEM usando a distância TOTAL percorrida
    (soma de todas as pernas da viagem) e o tipo de caminhão.
    Uma viagem com 3 paradas de 50 km cada = 150 km totais,
    não 3x a tarifa de 50 km.
    """
    if df_timeline.empty or not tabela_frete:
        return df_timeline, 0.0
    df = df_timeline.copy()

    caminhao_col = 'Caminhão' if 'Caminhão' in df.columns else 'Caminhao'

    # Distância total por viagem (soma das pernas)
    dist_total   = df.groupby('Viagem')['Distancia KM'].sum().to_dict()
    truck_viagem = df.groupby('Viagem')[caminhao_col].first().to_dict()

    frete_por_viagem = {}
    for trip, km_total in dist_total.items():
        tipo = str(truck_viagem.get(trip, ''))
        frete_por_viagem[trip] = calcular_frete_viagem(km_total, tipo, tabela_frete)

    # Distribui frete igualmente entre paradas (para fins de relatório)
    n_paradas = df.groupby('Viagem').size().to_dict()
    df["Frete (R$)"] = df["Viagem"].apply(
        lambda t: round(frete_por_viagem.get(t, 0) / n_paradas.get(t, 1), 2)
    )
    return df, round(sum(frete_por_viagem.values()), 2)

tabela_frete       = carregar_tabela_frete()

print("\n🧬 Algoritmo Genético — otimizando rotas...")
print(f"   Parâmetros: pop={AG_POPULACAO} | ger={AG_GERACOES} | "
      f"mut={AG_TAXA_MUTACAO:.0%} | penalidade=R${AG_PENALIDADE_ATR:.0f}/atraso")

import random
import copy

# Grupos de destino (reutiliza grupos_otm já construídos)
_ag_grupos = grupos_otm   # lista de dicts com cli, lat, lon, limite, engs
_ag_n      = len(_ag_grupos)

if _ag_n < 2:
    print("   ⚠️  Poucos destinos para o AG — pulando.")
else:
    # ── Decodificador: cromossomo → lista de trips ───────────
    def _ag_decodificar(cromossomo):
        """
        Converte permutação de índices em lista de trips.
        Agrupa destinos consecutivos até AG_MAX_PARADAS ou até
        o caminhão não comportar mais engradados.
        Retorna lista de grupos de destino por trip.
        """
        trips = []
        i = 0
        while i < len(cromossomo):
            trip_atual = []
            for j in range(i, min(i + AG_MAX_PARADAS, len(cromossomo))):
                trip_atual.append(_ag_grupos[cromossomo[j]])
                # Verifica se ainda cabe num caminhão
                truck_test, alocados_test, _ = empacotar_em_caminhao(
                    trip_atual, truck_list)
                if not alocados_test:
                    # Não cabe — reverte última parada
                    trip_atual = trip_atual[:-1]
                    break
            if not trip_atual:
                # Destino sozinho que não coube — coloca individualmente
                trip_atual = [_ag_grupos[cromossomo[i]]]
            trips.append(trip_atual)
            i += len(trip_atual)
        return trips

    # ── Fitness: calcula custo total de um cromossomo ─────────
    def _ag_fitness(cromossomo):
        """
        Calcula frete total + penalidade por atraso.
        Usa apenas Haversine (sem chamar API) para velocidade.
        Retorna (custo_total, n_trips, n_atrasos).
        """
        trips   = _ag_decodificar(cromossomo)
        h_base  = datetime.today().replace(
            hour=HORA_INICIO, minute=0, second=0, microsecond=0)
        custo   = 0.0
        atrasos = 0
        prox_doca  = h_base
        prox_patio = h_base

        for trip in trips:
            truck, alocados, _ = empacotar_em_caminhao(trip, truck_list)
            if not truck or not alocados:
                custo += AG_PENALIDADE_ATR * len(trip)
                atrasos += len(trip)
                continue

            # Carregamento paralelo Doca/Pátio
            t_carga = tempo_carregamento(sum(len(g['engs']) for g in trip))
            local_t = _prox_carga(truck['tipo'])
            if local_t == 'Pateo':
                h_saida    = prox_patio + timedelta(minutes=t_carga)
                prox_patio = h_saida
            else:
                h_saida   = prox_doca + timedelta(minutes=t_carga)
                prox_doca = h_saida

            # Simula a viagem usando apenas Haversine (rápido, sem API)
            p_ref = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
            h_cur = h_saida
            km_trip = 0.0
            for g in trip:
                dest = (g['lat'], g['lon'])
                km_seg, mins = _haversine(p_ref, dest)
                chegada = ajustar_pausas(h_cur, h_cur + timedelta(minutes=mins))
                if dt_min(chegada) > lim_min(g['limite']):
                    custo  += AG_PENALIDADE_ATR
                    atrasos += 1
                h_cur = chegada + timedelta(minutes=g.get('descarga', TEMPO_DESCARGA_MIN))
                p_ref = dest
                km_trip += km_seg

            # Frete da trip
            if tabela_frete:
                custo += calcular_frete_viagem(km_trip, truck['tipo'], tabela_frete)
            else:
                custo += km_trip * 3.5   # fallback R$/km genérico

        return custo, len(trips), atrasos

    # ── Geração da população inicial ─────────────────────────
    def _ag_pop_inicial(n_pop, n_genes):
        pop = []
        base = list(range(n_genes))
        # Primeiro indivíduo = ordem do plano otimizado (seed inteligente)
        ordem_otm = list(range(n_genes))
        pop.append(ordem_otm)
        for _ in range(n_pop - 1):
            ind = base.copy()
            random.shuffle(ind)
            pop.append(ind)
        return pop

    # ── Crossover OX (Order Crossover) ───────────────────────
    def _ag_crossover_ox(p1, p2):
        n   = len(p1)
        a, b = sorted(random.sample(range(n), 2))
        filho = [-1] * n
        filho[a:b+1] = p1[a:b+1]
        segmento = set(p1[a:b+1])
        pos = (b + 1) % n
        for gene in p2[b+1:] + p2[:b+1]:
            if gene not in segmento:
                filho[pos] = gene
                pos = (pos + 1) % n
        return filho

    # ── Mutação: troca 2-opt local ────────────────────────────
    def _ag_mutar(ind):
        n = len(ind)
        if n < 2:
            return ind
        a, b = sorted(random.sample(range(n), 2))
        ind[a], ind[b] = ind[b], ind[a]
        return ind

    # ── Seleção por torneio ───────────────────────────────────
    def _ag_torneio(pop, fitness_vals, k=3):
        candidatos = random.sample(range(len(pop)), k)
        return pop[min(candidatos, key=lambda i: fitness_vals[i])]

    # ── Loop evolutivo ────────────────────────────────────────
    random.seed(42)   # reprodutibilidade
    populacao    = _ag_pop_inicial(AG_POPULACAO, _ag_n)
    fitness_vals = [_ag_fitness(ind)[0] for ind in populacao]

    melhor_idx   = min(range(len(populacao)), key=lambda i: fitness_vals[i])
    melhor_ind   = populacao[melhor_idx][:]
    melhor_fit   = fitness_vals[melhor_idx]
    fit_inicial  = melhor_fit
    media_fit    = sum(fitness_vals) / len(fitness_vals)
    # Histórico de convergência: [(geracao, melhor, media), ...]
    historico_ag = [(0, melhor_fit, media_fit)]

    print(f"   Geração 0: melhor fitness = R${melhor_fit:,.0f}")

    for ger in range(1, AG_GERACOES + 1):
        # Elitismo: preserva os melhores
        ordem_rank = sorted(range(len(populacao)), key=lambda i: fitness_vals[i])
        nova_pop   = [populacao[i][:] for i in ordem_rank[:AG_ELITISMO]]

        while len(nova_pop) < AG_POPULACAO:
            pai1 = _ag_torneio(populacao, fitness_vals)
            if random.random() < AG_TAXA_CROSSOVER:
                pai2  = _ag_torneio(populacao, fitness_vals)
                filho = _ag_crossover_ox(pai1, pai2)
            else:
                filho = pai1[:]
            if random.random() < AG_TAXA_MUTACAO:
                filho = _ag_mutar(filho)
            nova_pop.append(filho)

        populacao    = nova_pop
        fitness_vals = [_ag_fitness(ind)[0] for ind in populacao]

        idx = min(range(len(populacao)), key=lambda i: fitness_vals[i])
        if fitness_vals[idx] < melhor_fit:
            melhor_fit = fitness_vals[idx]
            melhor_ind = populacao[idx][:]

        media_ger = sum(fitness_vals) / len(fitness_vals)
        historico_ag.append((ger, melhor_fit, media_ger))

        if ger % 20 == 0 or ger == AG_GERACOES:
            print(f"   Geração {ger:3d}: melhor = R${melhor_fit:,.0f}  "
                  f"(melhoria: {(fit_inicial-melhor_fit)/fit_inicial*100:+.1f}%)")

    # ── Gráfico de convergência HTML ─────────────────────────
    def _gerar_grafico_convergencia(historico, fit_ini, fit_fin, caminho_html):
        import json, webbrowser
        gers         = [h[0] for h in historico]
        melhor       = [h[1] for h in historico]
        media        = [h[2] for h in historico]
        n_geracoes   = len(historico)          # inclui geração 0
        total_planos = n_geracoes * AG_POPULACAO
        melhoria_pct = (fit_ini - fit_fin) / fit_ini * 100 if fit_ini > 0 else 0
        economia     = fit_ini - fit_fin
        hoje         = __import__('datetime').datetime.now().strftime('%d/%m/%Y %H:%M')

        # Planos acumulados por geração (para anotação no eixo X)
        planos_acum  = [g * AG_POPULACAO for g in gers]

        html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>AG — Convergência</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     background:#0f172a;color:#e2e8f0;padding:28px 32px}}
.header{{margin-bottom:8px}}
.header h1{{font-size:20px;font-weight:700;color:#f1f5f9;margin-bottom:4px}}
.header p {{font-size:13px;color:#64748b}}
.hero{{
  background:linear-gradient(135deg,#1D9E75 0%,#0F6E56 100%);
  border-radius:14px;padding:20px 28px;margin:20px 0;
  display:flex;align-items:center;gap:40px;flex-wrap:wrap
}}
.hero-item{{text-align:center}}
.hero-val{{font-size:36px;font-weight:800;color:white;line-height:1}}
.hero-val.red{{color:#fca5a5}}
.hero-sub{{font-size:11px;color:rgba(255,255,255,0.7);margin-top:4px;
           text-transform:uppercase;letter-spacing:.06em}}
.hero-div{{width:1px;height:56px;background:rgba(255,255,255,0.2)}}
.cards{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:24px}}
.card{{background:#1e293b;border-radius:12px;padding:14px 16px;border:1px solid #334155}}
.card-val{{font-size:20px;font-weight:700;color:#34d399;margin-bottom:3px}}
.card-val.amber{{color:#fbbf24}}
.card-val.blue {{color:#60a5fa}}
.card-val.red  {{color:#f87171}}
.card-lbl{{font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:.05em}}
.chart-wrap{{background:#1e293b;border-radius:14px;padding:24px;border:1px solid #334155;margin-bottom:16px}}
.chart-title{{font-size:14px;font-weight:600;color:#94a3b8;margin-bottom:16px}}
canvas{{max-height:380px}}
.footer{{text-align:center;font-size:11px;color:#334155;margin-top:8px}}
</style>
</head>
<body>

<div class="header">
  <h1>Algoritmo Genético — Aprendizado das Rotas</h1>
  <p>TTB Logistics &middot; {hoje} &middot; Suzano</p>
</div>

<div class="hero">
  <div class="hero-item">
    <div class="hero-val red">R$ {fit_ini:,.0f}</div>
    <div class="hero-sub">Custo inicial (pior rota)</div>
  </div>
  <div class="hero-div"></div>
  <div class="hero-item">
    <div class="hero-val">R$ {fit_fin:,.0f}</div>
    <div class="hero-sub">Melhor rota encontrada</div>
  </div>
  <div class="hero-div"></div>
  <div class="hero-item">
    <div class="hero-val">R$ {economia:,.0f}</div>
    <div class="hero-sub">Economia descoberta</div>
  </div>
  <div class="hero-div"></div>
  <div class="hero-item">
    <div class="hero-val">{melhoria_pct:.1f}%</div>
    <div class="hero-sub">Melhoria obtida</div>
  </div>
</div>

<div class="cards">
  <div class="card">
    <div class="card-val amber">{total_planos:,}</div>
    <div class="card-lbl">Planos avaliados</div>
  </div>
  <div class="card">
    <div class="card-val blue">{n_geracoes}</div>
    <div class="card-lbl">Gerações evolutivas</div>
  </div>
  <div class="card">
    <div class="card-val blue">{AG_POPULACAO}</div>
    <div class="card-lbl">Planos por geração</div>
  </div>
  <div class="card">
    <div class="card-val">{AG_ELITISMO}</div>
    <div class="card-lbl">Elites preservadas</div>
  </div>
  <div class="card">
    <div class="card-val red">{int(AG_TAXA_MUTACAO*100)}%</div>
    <div class="card-lbl">Taxa de mutação</div>
  </div>
</div>

<div class="chart-wrap">
  <div class="chart-title">Evolução do custo por geração
    <span style="font-size:11px;color:#475569;font-weight:400;margin-left:12px">
      Passe o mouse para ver detalhes &middot; Cada geração = {AG_POPULACAO} planos avaliados
    </span>
  </div>
  <canvas id="cv"></canvas>
</div>

<div class="footer">
  {total_planos:,} planos avaliados &middot; {n_geracoes} gerações &middot;
  Algoritmo: Genético com crossover OX + mutação 2-opt + elitismo
</div>

<script>
const GERS        = {json.dumps(gers)};
const MELHOR      = {json.dumps([round(v,2) for v in melhor])};
const MEDIA       = {json.dumps([round(v,2) for v in media])};
const PLANOS_ACUM = {json.dumps(planos_acum)};
const POP         = {AG_POPULACAO};

new Chart(document.getElementById('cv'), {{
  type: 'line',
  data: {{
    labels: GERS,
    datasets: [
      {{
        label: 'Melhor plano da geração',
        data: MELHOR,
        borderColor: '#34d399',
        backgroundColor: 'rgba(52,211,153,0.07)',
        borderWidth: 2.5,
        pointRadius: 0,
        pointHoverRadius: 6,
        pointHoverBackgroundColor: '#34d399',
        fill: true,
        tension: 0.35,
        order: 1
      }},
      {{
        label: 'Média da população',
        data: MEDIA,
        borderColor: '#60a5fa',
        backgroundColor: 'rgba(96,165,250,0.04)',
        borderWidth: 1.5,
        pointRadius: 0,
        pointHoverRadius: 5,
        pointHoverBackgroundColor: '#60a5fa',
        fill: true,
        tension: 0.35,
        borderDash: [5,3],
        order: 2
      }}
    ]
  }},
  options: {{
    responsive: true,
    interaction: {{ mode: 'index', intersect: false }},
    plugins: {{
      legend: {{
        labels: {{ color: '#94a3b8', font: {{ size: 12 }}, padding: 20 }}
      }},
      tooltip: {{
        backgroundColor: '#0f172a',
        borderColor: '#1D9E75',
        borderWidth: 1,
        titleColor: '#f1f5f9',
        bodyColor: '#94a3b8',
        padding: 12,
        callbacks: {{
          title: items => {{
            const g = items[0].label;
            const acum = PLANOS_ACUM[items[0].dataIndex];
            return 'Geração ' + g + '  ·  ' + acum.toLocaleString('pt-BR') + ' planos avaliados';
          }},
          label: ctx => {{
            const v = ctx.parsed.y;
            return ' ' + ctx.dataset.label + ': R$ ' +
                   v.toLocaleString('pt-BR', {{minimumFractionDigits:0, maximumFractionDigits:0}});
          }},
          afterBody: items => {{
            const i  = items[0].dataIndex;
            const mb = MELHOR[i], mm = MEDIA[i];
            const gap = ((mm - mb) / mb * 100).toFixed(1);
            return ['', ' Dispersão pop.: +' + gap + '% acima do melhor'];
          }}
        }}
      }}
    }},
    scales: {{
      x: {{
        title: {{ display: true, text: 'Geração', color: '#475569', font:{{size:12}} }},
        ticks: {{
          color: '#475569',
          maxTicksLimit: 16,
          callback: (v, i) => GERS[i]
        }},
        grid: {{ color: 'rgba(255,255,255,0.03)' }}
      }},
      y: {{
        title: {{ display: true, text: 'Custo da rota (R$)', color: '#475569', font:{{size:12}} }},
        ticks: {{
          color: '#475569',
          callback: v => 'R$ ' + (v/1000).toFixed(0) + 'k'
        }},
        grid: {{ color: 'rgba(255,255,255,0.05)' }}
      }}
    }}
  }}
}});
</script>
</body>
</html>"""

        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(html)
        salvar_e_abrir(caminho_html, silent=True)
        print(f"   📈 Gráfico de convergência → {__import__('os').path.basename(caminho_html)}")

    _caminho_conv = os.path.join(DRIVE_PATH, 'AG_convergencia.html')
    _gerar_grafico_convergencia(historico_ag, fit_inicial, melhor_fit, _caminho_conv)

    # ── Decodifica o melhor indivíduo em timeline ─────────────
    print(f"\n   🏆 Melhor solução AG: R${melhor_fit:,.0f}  "
          f"(era R${fit_inicial:,.0f} no otimizado heurístico)")

    trips_ag     = _ag_decodificar(melhor_ind)
    timeline_ag  = []
    id_trip_ag   = 1
    _h_ag        = datetime.today().replace(
        hour=HORA_INICIO, minute=0, second=0, microsecond=0)
    prox_doca_ag  = _h_ag
    prox_patio_ag = _h_ag

    for trip_g in trips_ag:
        truck_ag, alocados_ag, sobram_ag = empacotar_em_caminhao(trip_g, truck_list)
        if not truck_ag or not alocados_ag:
            continue

        qtd_engs_ag = sum(len(g['engs']) for g in trip_g)
        t_carga_ag  = tempo_carregamento(qtd_engs_ag)
        local_ag    = _prox_carga(truck_ag['tipo'])
        if local_ag == 'Pateo':
            h_ini_ag       = prox_patio_ag
            h_sai_ag       = h_ini_ag + timedelta(minutes=t_carga_ag)
            prox_patio_ag  = h_sai_ag
        else:
            h_ini_ag       = prox_doca_ag
            h_sai_ag       = h_ini_ag + timedelta(minutes=t_carga_ag)
            prox_doca_ag   = h_sai_ag

        # Simula com rotas do cache (sem nova chamada à API)
        rotas_ag, _, chegadas_ag = simular_viagem(
            trip_g, mem_cache, gmaps, AG_USAR_CACHE_API, controlador_api,
            h_inicio=h_sai_ag)

        vol_truck_ag   = truck_ag['vol']
        cli_map_ag     = {}
        for eng in alocados_ag:
            cli_map_ag.setdefault(eng['cli'], []).append(eng)

        vol_trip_ag = min(
            sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                if 'dim_f' in e else e.get('vol', 0)
                for e in alocados_ag),
            vol_truck_ag)
        ocup_ag = round(vol_trip_ag / vol_truck_ag * 100, 1)

        p_ref_ag = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
        h_ref_ag = h_sai_ag
        trip_id_str = f"AG_{id_trip_ag:03d}"

        for g in trip_g:
            cli  = g['cli']
            engs = cli_map_ag.get(cli, [])
            if not engs:
                continue
            dest   = (g['lat'], g['lon'])
            rid    = _rota_id(p_ref_ag, dest)
            km, _m = rotas_ag.get(rid, _haversine(p_ref_ag, dest))
            chegada_bruta = h_ref_ag + timedelta(minutes=_m)
            chegada = ajustar_pausas(h_ref_ag, chegada_bruta)
            no_prazo = dt_min(chegada) <= lim_min(g['limite'])

            timeline_ag.append({
                'Viagem':              trip_id_str,
                'Caminhão':            truck_ag['tipo'],
                'Local Carga':         local_ag,
                'Início Carga':        h_ini_ag.strftime("%H:%M"),
                'Tempo Carga (min)':   t_carga_ag,
                'Saída Armazém':       h_sai_ag.strftime("%H:%M"),
                'Qtd Engradados':      len(engs),
                'Cliente':             cli,
                'NFs':                 "/".join(sorted(set(e['nf'] for e in engs))),
                'Distancia KM':        round(km, 2),
                'Chegada':             chegada.strftime("%H:%M"),
                'Tempo Descarga (min)': g.get('descarga', TEMPO_DESCARGA_MIN),
                'Limite Entrega':      g['limite'],
                'Dentro do Prazo':     '✅' if no_prazo else '❌',
                'Ocupação %':          ocup_ag,
                'Fonte Rota':          'Cache/Haversine'
            })

            h_ref_ag = chegada + timedelta(minutes=g.get('descarga', TEMPO_DESCARGA_MIN))
            p_ref_ag = dest

        id_trip_ag += 1

    df_ag = pd.DataFrame(timeline_ag)
    if not df_ag.empty:
        df_ag, co2_ag_kg, co2_ag_t = calcular_co2(df_ag, consumo_combustivel)
        salvar_excel(df_ag, os.path.join(DRIVE_PATH, '06_TIMELINE_AG.xlsx'),
                     '06_TIMELINE_AG.xlsx')
        n_viag_ag   = df_ag['Viagem'].nunique()
        no_prazo_ag = int((df_ag['Dentro do Prazo'] == '✅').sum())
        print(f"   ✅ {n_viag_ag} viagem(ns) AG → 06_TIMELINE_AG.xlsx")
        print(f"   ⏱️  Entregas no prazo : {no_prazo_ag}/{len(df_ag)}")
        print(f"   🚛 Caminhões usados   : {n_viag_ag}")
        print(f"   🚛 Ocupação média     : {df_ag['Ocupação %'].mean():.1f}%")
        if co2_ag_kg > 0:
            print(f"   🌿 Emissão CO2 est.  : {co2_ag_kg:.1f} kg ({co2_ag_t:.3f} t)")

        # Adiciona AG no comparativo
        dist_ag_km = df_ag['Distancia KM'].sum() if 'Distancia KM' in df_ag.columns else 0
        df_ag, frete_ag = calcular_frete_total(df_ag, tabela_frete)
        _dist_otm_km_ag  = df_otm['Distancia KM'].sum()       if not df_otm.empty else 0
        _dist_real_km_ag = df_tl['Distancia KM'].sum()        if not df_tl.empty else 0
        _no_prazo_real_ag= int((df_tl['Dentro do Prazo'] == '✅').sum()) if not df_tl.empty else 0
        _no_prazo_otm_ag = int((df_otm['Dentro do Prazo'] == '✅').sum()) if not df_otm.empty else 0
        _n_cam_real_ag   = df_tl['Viagem'].nunique()           if not df_tl.empty else 0
        _n_cam_otm_ag    = df_otm['Viagem'].nunique()          if not df_otm.empty else 0
        _co2_otm_ag      = co2_otm_kg  if 'co2_otm_kg'  in dir() else 0
        _co2_real_ag     = co2_real_kg if 'co2_real_kg' in dir() else 0
        print(f"\n   ── Comparativo rápido ──────────────────────────────────")
        print(f"   {'Plano':<18} {'Trips':>6} {'Prazo':>8} {'KM':>8} {'CO2':>8}")
        print(f"   {'─'*52}")
        print(f"   {'Real':<18} {_n_cam_real_ag:>6} "
              f"{_no_prazo_real_ag}/{len(df_tl):>3} "
              f"{_dist_real_km_ag:>8.0f} {_co2_real_ag:>7.0f}kg")
        print(f"   {'Otimizado':<18} {_n_cam_otm_ag:>6} "
              f"{_no_prazo_otm_ag}/{len(df_otm):>3} "
              f"{_dist_otm_km_ag:>8.0f} {_co2_otm_ag:>7.0f}kg")
        print(f"   {'AG':<18} {n_viag_ag:>6} "
              f"{no_prazo_ag}/{len(df_ag):>3} "
              f"{dist_ag_km:>8.0f} {co2_ag_kg:>7.0f}kg")
        print(f"   {'─'*52}")
    else:
        print("   ⚠️  AG não gerou viagens — verifique parâmetros.")

viagens_otm         = id_trip_otm - 1
no_prazo_otm        = int((df_otm['Dentro do Prazo'] == '✅').sum()) if not df_otm.empty else 0
n_caminhoes_otm     = df_otm['Viagem'].nunique() if not df_otm.empty else 0
cli_por_viagem_otm  = df_otm.groupby('Viagem')['Cliente'].count().mean() if not df_otm.empty else 0.0

print(f"   ✅ {viagens_otm} viagem(ns) → 06_TIMELINE_OTIMIZADO.xlsx")
if not df_otm.empty:
    print(f"   👥 Clientes/viagem     : {cli_por_viagem_otm:.1f} "
          f"(original: {cli_por_viagem_real:.1f})")
    print(f"   ⏱️  Entregas no prazo  : {no_prazo_otm}/{len(df_otm)}")
    print(f"   🚛 Ocupação média      : {df_otm['Ocupação %'].mean():.1f}%")
    if co2_otm_kg > 0:
        print(f"   🌿 Emissão CO2 est.   : {co2_otm_kg:.1f} kg ({co2_otm_t:.3f} t)")
else:
    print(f"   ⚠️  Nenhuma viagem gerada no otimizado — verifique eng_final e clusters_otm")
    print(f"      eng_final: {len(eng_final)} engradados | clusters_otm: {len(clusters_otm)} clusters")

# Distâncias totais por plano (calculadas aqui para o comparativo)
dist_real_km = df_tl['Distancia KM'].sum() if not df_tl.empty else 0
dist_otm_km  = df_otm['Distancia KM'].sum() if not df_otm.empty else 0

# Comparativo entre planos — exibe após cálculo do frete (referenciado abaixo)
# Variáveis de frete serão preenchidas na seção 13B; inicializa com 0
_frete_real_comp = 0.0
_frete_otm_comp  = 0.0

def _imprimir_comparativo(fr_real=0.0, fr_otm=0.0, fr_ag=None):
    # Dados AG
    _df_ag_ok = 'df_ag' in globals() and isinstance(df_ag, pd.DataFrame) and not df_ag.empty
    tem_ag    = _df_ag_ok
    df_ag_c   = df_ag if tem_ag else pd.DataFrame()
    co2_ag_c  = (co2_ag_kg if ('co2_ag_kg' in globals() and co2_ag_kg > 0) else 0.0)
    if fr_ag is None and tem_ag:
        fr_ag = globals().get('frete_ag', 0.0)

    # Dados DBSCAN
    _df_db_ok  = 'df_dbscan' in globals() and isinstance(df_dbscan, pd.DataFrame) and not df_dbscan.empty
    tem_db     = _df_db_ok
    df_db_c    = df_dbscan    if tem_db  else pd.DataFrame()
    co2_db_c   = globals().get('co2_dbscan_kg', 0.0)   if tem_db  else 0.0
    fr_db      = globals().get('frete_dbscan',  0.0)   if tem_db  else 0.0

    # Dados DBSCAN+AG
    _df_dba_ok = 'df_dbscan_ag' in globals() and isinstance(df_dbscan_ag, pd.DataFrame) and not df_dbscan_ag.empty
    tem_dba    = _df_dba_ok
    df_dba_c   = df_dbscan_ag if tem_dba else pd.DataFrame()
    co2_dba_c  = globals().get('co2_dbscan_ag_kg', 0.0) if tem_dba else 0.0
    fr_dba     = globals().get('frete_dbscan_ag',  0.0) if tem_dba else 0.0

    # Dados RL
    _df_rl_ok  = 'df_rl' in globals() and isinstance(df_rl, pd.DataFrame) and not df_rl.empty
    tem_rl     = _df_rl_ok
    df_rl_c    = df_rl        if tem_rl  else pd.DataFrame()
    co2_rl_c   = globals().get('co2_rl_kg',  0.0) if tem_rl  else 0.0
    fr_rl      = globals().get('frete_rl',   0.0) if tem_rl  else 0.0
    def _ocup(df): return df['Ocupação %'].mean()   if not df.empty and 'Ocupação %'   in df.columns else 0
    def _n(df):    return df['Viagem'].nunique()     if not df.empty else 0
    def _cli(df):  return round(df.groupby('Viagem')['Cliente'].count().mean(), 1) if not df.empty else 0
    def _dist(df): return df['Distancia KM'].sum()  if not df.empty and 'Distancia KM' in df.columns else 0

    ocup_r = _ocup(df_tl);   ocup_o = _ocup(df_otm)
    ocup_a = _ocup(df_ag_c); ocup_d = _ocup(df_db_c); ocup_da = _ocup(df_dba_c); ocup_rl = _ocup(df_rl_c)

    n_r = _n(df_tl);  n_o = _n(df_otm)
    n_a = _n(df_ag_c); n_d = _n(df_db_c); n_da = _n(df_dba_c); n_rl = _n(df_rl_c)

    cli_r = _cli(df_tl);  cli_o = _cli(df_otm)
    cli_a = _cli(df_ag_c); cli_d = _cli(df_db_c); cli_da = _cli(df_dba_c); cli_rl = _cli(df_rl_c)

    dist_r = _dist(df_tl);   dist_o = _dist(df_otm)
    dist_a = _dist(df_ag_c); dist_d = _dist(df_db_c); dist_da = _dist(df_dba_c); dist_rl = _dist(df_rl_c)

    def _co2tipo(df):
        col = 'Caminhão' if 'Caminhão' in df.columns else 'Caminhao'
        if df.empty or 'CO2 (kg)' not in df.columns: return {}
        return df.groupby(col)['CO2 (kg)'].sum().to_dict()

    def _tipos(df):
        col = 'Caminhão' if 'Caminhão' in df.columns else 'Caminhao'
        if df.empty or col not in df.columns: return {}
        return df.groupby('Viagem')[col].first().value_counts().to_dict()

    tipos_r  = _tipos(df_tl);    tipos_o  = _tipos(df_otm)
    tipos_a  = _tipos(df_ag_c);  tipos_d  = _tipos(df_db_c)
    tipos_da = _tipos(df_dba_c); tipos_rl = _tipos(df_rl_c)
    all_tipos = sorted(set(list(tipos_r)+list(tipos_o)+list(tipos_a)+list(tipos_d)+list(tipos_da)+list(tipos_rl)))

    co2_r = globals().get('co2_real_kg', 0); co2_o = globals().get('co2_otm_kg', 0)

    co2t_r = _co2tipo(df_tl);    co2t_o  = _co2tipo(df_otm)
    co2t_a = _co2tipo(df_ag_c);  co2t_d  = _co2tipo(df_db_c)
    co2t_da= _co2tipo(df_dba_c); co2t_rl = _co2tipo(df_rl_c)
    all_co2t = sorted(set(list(co2t_r)+list(co2t_o)+list(co2t_a)+list(co2t_d)+list(co2t_da)+list(co2t_rl)))

    # ── Larguras de coluna ────────────────────────────────────
    W  = 28    # label
    C  = 9     # cada coluna de valor

    ncols = 2 + (1 if tem_ag else 0) + (1 if tem_db else 0) + (1 if tem_dba else 0) + (1 if tem_rl else 0)
    sep   = '─' * (W + C * ncols + ncols)

    def _hdr():
        h = f"   {'Métrica':<{W}} {'Original':>{C}} {'Otimizado':>{C}}"
        if tem_ag:  h += f" {'AG':>{C}}"
        if tem_db:  h += f" {'DBSCAN':>{C}}"
        if tem_dba: h += f" {'DBSCAN+AG':>{C}}"
        if tem_rl:  h += f" {'RL(DQN)':>{C}}"
        return h

    def _lin(label, vr, vo, va='', vd='', vda='', vrl='', fmt='>9'):
        s = f"   {label:<{W}} {vr:{fmt}} {vo:{fmt}}"
        if tem_ag:  s += f" {va  if va  != '' else '':>{C}}"
        if tem_db:  s += f" {vd  if vd  != '' else '':>{C}}"
        if tem_dba: s += f" {vda if vda != '' else '':>{C}}"
        if tem_rl:  s += f" {vrl if vrl != '' else '':>{C}}"
        return s

    def _pct(base, novo):
        if base <= 0 or novo == '': return ''
        try: return f"{(base-float(str(novo).replace(',','')))/base*100:+.1f}%"
        except: return ''

    titulo = ("📊 COMPARATIVO — 6 PLANOS" if tem_rl else "📊 COMPARATIVO — 5 PLANOS")
    titulo += " (Original | Otimizado | AG | DBSCAN | DBSCAN+AG" + (" | RL-DQN)" if tem_rl else ")")
    print(f"\n   {sep}")
    print(f"   {titulo}")
    print(f"   {sep}")
    print(_hdr())
    print(f"   {sep}")

    # Viagens
    print(_lin('Viagens (caminhões)', n_r, n_o,
               n_a if tem_ag else '', n_d if tem_db else '',
               n_da if tem_dba else '', n_rl if tem_rl else ''))
    for tp in all_tipos:
        print(_lin(f'  └ {tp}',
                   tipos_r.get(tp,0),  tipos_o.get(tp,0),
                   tipos_a.get(tp,0)   if tem_ag  else '',
                   tipos_d.get(tp,0)   if tem_db  else '',
                   tipos_da.get(tp,0)  if tem_dba else '',
                   tipos_rl.get(tp,0)  if tem_rl  else ''))
    print(f"   {sep}")

    # Métricas
    print(_lin('Clientes/viagem',    f'{cli_r:.1f}', f'{cli_o:.1f}',
               f'{cli_a:.1f}'  if tem_ag  else '', f'{cli_d:.1f}'  if tem_db  else '',
               f'{cli_da:.1f}' if tem_dba else '', f'{cli_rl:.1f}' if tem_rl  else '', fmt='>9'))
    print(_lin('Ocupação média (%)', f'{ocup_r:.1f}', f'{ocup_o:.1f}',
               f'{ocup_a:.1f}'  if tem_ag  else '', f'{ocup_d:.1f}'  if tem_db  else '',
               f'{ocup_da:.1f}' if tem_dba else '', f'{ocup_rl:.1f}' if tem_rl  else '', fmt='>9'))
    print(_lin('Distância total (km)', f'{dist_r:.1f}', f'{dist_o:.1f}',
               f'{dist_a:.1f}'  if tem_ag  else '', f'{dist_d:.1f}'  if tem_db  else '',
               f'{dist_da:.1f}' if tem_dba else '', f'{dist_rl:.1f}' if tem_rl  else '', fmt='>9'))

    # Redução distância
    s_rd = f"   {'Redução distância (%)':<{W}} {'':>{C}} {_pct(dist_r, dist_o):>{C}}"
    if tem_ag:  s_rd += f" {_pct(dist_r, dist_a):>{C}}"
    if tem_db:  s_rd += f" {_pct(dist_r, dist_d):>{C}}"
    if tem_dba: s_rd += f" {_pct(dist_r, dist_da):>{C}}"
    if tem_rl:  s_rd += f" {_pct(dist_r, dist_rl):>{C}}"
    print(s_rd)

    if fr_real > 0:
        print(_lin('Frete total (R$)', f'{fr_real:,.0f}', f'{fr_otm:,.0f}',
                   f'{fr_ag:,.0f}'  if tem_ag  else '', f'{fr_db:,.0f}'  if tem_db  else '',
                   f'{fr_dba:,.0f}' if tem_dba else '', f'{fr_rl:,.0f}'  if tem_rl  else '', fmt='>9'))
        s_pf = f"   {'Redução frete (%)':<{W}} {'':>{C}} {_pct(fr_real, fr_otm):>{C}}"
        if tem_ag:  s_pf += f" {_pct(fr_real, fr_ag):>{C}}"
        if tem_db:  s_pf += f" {_pct(fr_real, fr_db):>{C}}"
        if tem_dba: s_pf += f" {_pct(fr_real, fr_dba):>{C}}"
        if tem_rl:  s_pf += f" {_pct(fr_real, fr_rl):>{C}}"
        print(s_pf)

    if co2_r > 0 or co2_o > 0:
        print(f"   {sep}")
        print(_lin('Emissão CO₂ (kg)', f'{co2_r:.0f}', f'{co2_o:.0f}',
                   f'{co2_ag_c:.0f}'  if tem_ag  else '', f'{co2_db_c:.0f}'  if tem_db  else '',
                   f'{co2_dba_c:.0f}' if tem_dba else '', f'{co2_rl_c:.0f}'  if tem_rl  else '', fmt='>9'))
        s_co2 = f"   {'Redução CO₂ (%)':<{W}} {'':>{C}} {_pct(co2_r, co2_o):>{C}}"
        if tem_ag:  s_co2 += f" {_pct(co2_r, co2_ag_c):>{C}}"
        if tem_db:  s_co2 += f" {_pct(co2_r, co2_db_c):>{C}}"
        if tem_dba: s_co2 += f" {_pct(co2_r, co2_dba_c):>{C}}"
        if tem_rl:  s_co2 += f" {_pct(co2_r, co2_rl_c):>{C}}"
        print(s_co2)
        for tp in all_co2t:
            print(_lin(f'  └ {tp}',
                       f'{co2t_r.get(tp,0):.0f}',  f'{co2t_o.get(tp,0):.0f}',
                       f'{co2t_a.get(tp,0):.0f}'   if tem_ag  else '',
                       f'{co2t_d.get(tp,0):.0f}'   if tem_db  else '',
                       f'{co2t_da.get(tp,0):.0f}'  if tem_dba else '',
                       f'{co2t_rl.get(tp,0):.0f}'  if tem_rl  else '', fmt='>9'))

    print(f"   {sep}")
    print(f"   💰 Frete original   : R$ {fr_real:,.2f}")
    print(f"   💰 Frete otimizado  : R$ {fr_otm:,.2f}")
    if tem_ag:  print(f"   💰 Frete AG         : R$ {fr_ag:,.2f}")
    if tem_db:  print(f"   💰 Frete DBSCAN     : R$ {fr_db:,.2f}")
    if tem_dba: print(f"   💰 Frete DBSCAN+AG  : R$ {fr_dba:,.2f}")
    if tem_rl:  print(f"   💰 Frete RL (DQN)   : R$ {fr_rl:,.2f}")

    ocup_r   = df_tl['Ocupação %'].mean()   if not df_tl.empty   else 0
    ocup_o   = df_otm['Ocupação %'].mean()  if not df_otm.empty  else 0
    ocup_a   = df_ag_c['Ocupação %'].mean() if not df_ag_c.empty else 0

    n_cam_r  = n_caminhoes_real
    n_cam_o  = n_caminhoes_otm
    n_cam_a  = df_ag_c['Viagem'].nunique() if not df_ag_c.empty else 0

    cli_ag   = df_ag_c.groupby('Viagem')['Cliente'].count().mean() if not df_ag_c.empty else 0.0
    dist_ag  = df_ag_c['Distancia KM'].sum() if not df_ag_c.empty else 0.0

    red_dist_o  = dist_real_km - dist_otm_km
    pct_dist_o  = red_dist_o / dist_real_km * 100 if dist_real_km > 0 else 0
    red_dist_a  = dist_real_km - dist_ag
    pct_dist_a  = red_dist_a / dist_real_km * 100 if dist_real_km > 0 else 0

    red_fr_o = fr_real - fr_otm
    pct_fr_o = red_fr_o / fr_real * 100 if fr_real > 0 else 0
    red_fr_a = fr_real - (fr_ag or 0)
    pct_fr_a = red_fr_a / fr_real * 100 if fr_real > 0 else 0

    # Largura de colunas
    W = 30   # métrica
    C = 9    # cada coluna de valor

    sep = '─' * (W + C*3 + (C if tem_ag else 0) + 6)

    def _linha(label, vr, vo, va=None, fmt='>9', sinal=False):
        s = f"   {label:<{W}} {vr:{fmt}}"
        s += f" {vo:{fmt}}"
        if tem_ag:
            s += f" {va if va is not None else '':>{C}}"
        return s

    print(f"\n   {sep}")
    titulo = "📊 COMPARATIVO ORIGINAL vs OTIMIZADO" + (" vs AG" if tem_ag else "")
    print(f"   {titulo}")
    print(f"   {sep}")

    # Cabeçalho
    hdr = f"   {'Métrica':<{W}} {'Original':>{C}} {'Otimizado':>{C}}"
    if tem_ag: hdr += f" {'AG':>{C}}"
    print(hdr)
    print(f"   {sep}")

    # ── Viagens por tipo ─────────────────────────────────────
    def _contar_tipos(df):
        col = 'Caminhão' if 'Caminhão' in df.columns else 'Caminhao'
        if df.empty or col not in df.columns: return {}
        return df.groupby('Viagem')[col].first().value_counts().to_dict()

    tipos_r  = _contar_tipos(df_tl)
    tipos_o  = _contar_tipos(df_otm)
    tipos_a  = _contar_tipos(df_ag_c)
    all_tipos = sorted(set(list(tipos_r.keys()) + list(tipos_o.keys()) + list(tipos_a.keys())))

    linha_cam = f"   {'Viagens (caminhoes)':<{W}} {n_cam_r:>{C}} {n_cam_o:>{C}}"
    if tem_ag: linha_cam += f" {n_cam_a:>{C}}"
    print(linha_cam)

    for _tp in all_tipos:
        _qr = tipos_r.get(_tp, 0)
        _qo = tipos_o.get(_tp, 0)
        _qa = tipos_a.get(_tp, 0)
        linha_tp = f"   {'  └ ' + str(_tp):<{W}} {_qr:>{C}} {_qo:>{C}}"
        if tem_ag: linha_tp += f" {_qa:>{C}}"
        print(linha_tp)

    print(f"   {sep}")

    # ── Métricas principais ──────────────────────────────────
    def _pr(label, vr, vo, va, fmt='9.1f'):
        s = f"   {label:<{W}} {vr:>{fmt}} {vo:>{fmt}}"
        if tem_ag: s += f" {va:>{fmt}}"
        return s

    print(_pr('Clientes/viagem',    cli_por_viagem_real, cli_por_viagem_otm, cli_ag))
    print(_pr('Ocupação média (%)', ocup_r,              ocup_o,             ocup_a))
    print(_pr('Distância total (km)', dist_real_km, dist_otm_km, dist_ag, fmt='9.1f'))

    # Redução distância
    s_rd = f"   {'Redução distância (%)':<{W}} {'':>{C}} {pct_dist_o:>+8.1f}%"
    if tem_ag: s_rd += f" {pct_dist_a:>+8.1f}%"
    print(s_rd)

    if fr_real > 0 or fr_otm > 0:
        s_fr = f"   {'Frete total (R$)':<{W}} {fr_real:>{C},.0f} {fr_otm:>{C},.0f}"
        if tem_ag: s_fr += f" {fr_ag:>{C},.0f}"
        print(s_fr)
        s_pf = f"   {'Redução frete (%)':<{W}} {'':>{C}} {pct_fr_o:>+8.1f}%"
        if tem_ag: s_pf += f" {pct_fr_a:>+8.1f}%"
        print(s_pf)

    if co2_real_kg > 0 or co2_otm_kg > 0 or co2_ag_c > 0:
        red_co2_o = co2_real_kg - co2_otm_kg
        pct_co2_o = red_co2_o / co2_real_kg * 100 if co2_real_kg > 0 else 0
        red_co2_a = co2_real_kg - co2_ag_c
        pct_co2_a = red_co2_a / co2_real_kg * 100 if co2_real_kg > 0 else 0

        print(f"   {sep}")

        s_co2 = f"   {'Emissão CO2 (kg)':<{W}} {co2_real_kg:>{C},.0f} {co2_otm_kg:>{C},.0f}"
        if tem_ag: s_co2 += f" {co2_ag_c:>{C},.0f}"
        print(s_co2)

        s_pco2 = f"   {'Redução CO2 (%)':<{W}} {'':>{C}} {pct_co2_o:>+8.1f}%"
        if tem_ag: s_pco2 += f" {pct_co2_a:>+8.1f}%"
        print(s_pco2)

        # CO2 por tipo de caminhão
        col_c = 'Caminhão' if 'Caminhão' in df_tl.columns else 'Caminhao'
        co2_r = df_tl.groupby(col_c)['CO2 (kg)'].sum() if 'CO2 (kg)' in df_tl.columns else pd.Series()
        co2_o = df_otm.groupby(col_c)['CO2 (kg)'].sum() if 'CO2 (kg)' in df_otm.columns else pd.Series()
        co2_a = df_ag_c.groupby(col_c)['CO2 (kg)'].sum() if 'CO2 (kg)' in df_ag_c.columns else pd.Series()
        all_tp_co2 = sorted(set(list(co2_r.index) + list(co2_o.index) + list(co2_a.index)))
        for _tp in all_tp_co2:
            s = f"   {'  └ ' + str(_tp):<{W}} {co2_r.get(_tp,0):>{C},.0f} {co2_o.get(_tp,0):>{C},.0f}"
            if tem_ag: s += f" {co2_a.get(_tp,0):>{C},.0f}"
            print(s)

    print(f"   {sep}")





# =============================================================
# SEÇÃO 13D — PLANO DBSCAN (clustering geográfico puro)
# Referenciado no TCC como "K-Means Geográfico"
#
# Estratégia:
#   1. DBSCAN agrupa destinos por proximidade geográfica (raio RAIO_DBSCAN_KM)
#   2. Clusters grandes (> 1 caminhão) são subdivididos recursivamente
#   3. Cada cluster vira uma viagem — sem otimização de ordem/frete
#   4. Permite comparar clustering puro vs AG vs heurística greedy
# =============================================================
print("\n📍 Plano DBSCAN — clustering geográfico...")

df_dbscan    = pd.DataFrame()
timeline_dbscan  = []
co2_dbscan_kg    = 0.0
frete_dbscan     = 0.0

try:
    from sklearn.cluster import DBSCAN as _DBSCAN
    from sklearn.preprocessing import StandardScaler as _StandardScaler
    import math as _math

    # Coordenadas dos grupos de destino (reutiliza grupos_otm já construídos)
    _db_grupos = construir_grupos_destino(eng_final)
    for g in _db_grupos:
        g['limite'] = limite_otimizado_cliente(g['cli'])

    if len(_db_grupos) < 2:
        print("   ⚠️  Poucos destinos para DBSCAN — pulando.")
    else:
        # Converte raio km → radianos para DBSCAN com métrica haversine
        _coords = np.array([[g['lat'], g['lon']] for g in _db_grupos])
        _coords_rad = np.radians(_coords)
        _eps_rad = RAIO_DBSCAN_KM / 6371.0   # raio Terra em km

        _db = _DBSCAN(eps=_eps_rad, min_samples=1, algorithm='ball_tree',
                      metric='haversine').fit(_coords_rad)
        _labels = _db.labels_

        # Agrupa destinos por cluster
        _clusters_db = {}
        for i, lbl in enumerate(_labels):
            _clusters_db.setdefault(lbl, []).append(_db_grupos[i])

        print(f"   📍 DBSCAN: {len(_clusters_db)} cluster(s) "
              f"com raio ≤ {RAIO_DBSCAN_KM} km "
              f"({len(_db_grupos)} destinos)")

        # Ordena clusters por urgência (limite mais restritivo primeiro)
        _clusters_db_list = sorted(
            _clusters_db.values(),
            key=lambda c: min(lim_min(g['limite']) for g in c)
        )

        id_trip_db       = 1
        _prox_doca_db    = datetime.today().replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
        _prox_patio_db   = _prox_doca_db

        for _cl in _clusters_db_list:
            # Subdivide clusters que não cabem num único caminhão
            _pendentes = [_cl]
            _iter_seg  = 0
            while _pendentes and _iter_seg < 20:
                _iter_seg += 1
                _grupo = _pendentes.pop(0)
                _paradas = sorted(_grupo, key=lambda g: lim_min(g['limite']))

                _truck_db, _aloc_db, _sob_db = empacotar_em_caminhao(_paradas, truck_list)
                if not _truck_db or not _aloc_db:
                    continue

                if _sob_db:
                    # Sobras voltam como novo grupo
                    _clis_sob = {e['cli'] for e in _sob_db}
                    _paradas_sob = [g for g in _paradas if g['cli'] in _clis_sob]
                    _paradas     = [g for g in _paradas if g['cli'] not in _clis_sob]
                    if _paradas_sob:
                        _pendentes.append(_paradas_sob)
                    if not _paradas:
                        continue

                # Carregamento paralelo Doca/Pátio
                _qtd_e_db = sum(len(g['engs']) for g in _paradas)
                _t_c_db   = tempo_carregamento(_qtd_e_db)
                _loc_db   = _prox_carga(_truck_db['tipo'])
                if _loc_db == 'Pateo':
                    _h_ini_db    = _prox_patio_db
                    _h_sai_db    = _h_ini_db + timedelta(minutes=_t_c_db)
                    _prox_patio_db = _h_sai_db
                else:
                    _h_ini_db   = _prox_doca_db
                    _h_sai_db   = _h_ini_db + timedelta(minutes=_t_c_db)
                    _prox_doca_db = _h_sai_db

                _rotas_db, _, _cheg_db = simular_viagem(
                    _paradas, mem_cache, gmaps, API_DISPONIVEL, controlador_api,
                    h_inicio=_h_sai_db)

                _vol_db = min(
                    sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                        if 'dim_f' in e else e.get('vol', 0)
                        for e in _aloc_db),
                    _truck_db['vol'])

                _cli_map_db = {}
                for _e in _aloc_db:
                    _cli_map_db.setdefault(_e['cli'], []).append(_e)

                _p_ref_db = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
                _h_ref_db = _h_sai_db
                _trip_db  = f"DB_{id_trip_db:03d}"

                for _g in _paradas:
                    _cli  = _g['cli']
                    _engs = _cli_map_db.get(_cli, [])
                    if not _engs:
                        continue
                    _dest    = (_g['lat'], _g['lon'])
                    _rid     = _rota_id(_p_ref_db, _dest)
                    _km, _m  = _rotas_db.get(_rid, _haversine(_p_ref_db, _dest))
                    _cheg_b  = _h_ref_db + timedelta(minutes=_m)
                    _cheg    = ajustar_pausas(_h_ref_db, _cheg_b)
                    _ok      = dt_min(_cheg) <= lim_min(_g['limite'])

                    timeline_dbscan.append({
                        'Viagem':              _trip_db,
                        'Caminhão':            _truck_db['tipo'],
                        'Vol Caminhão (m³)':   round(_truck_db['vol'], 2),
                        'Vol Carga Total (m³)': round(_vol_db, 4),
                        'Ocupação %':          round(_vol_db / _truck_db['vol'] * 100, 1),
                        'Cliente':             _cli,
                        'Qtd Engradados':      len(_engs),
                        'Tipos Engradados':    ", ".join(sorted(set(e['tipo'] for e in _engs))),
                        'NFs':                 "/".join(sorted(set(e['nf'] for e in _engs))),
                        'Início Carga':        _h_ini_db.strftime("%H:%M"),
                        'Local Carga':         _loc_db,
                        'Tempo Carga (min)':   _t_c_db,
                        'Saída Armazém':       _h_sai_db.strftime("%H:%M"),
                        'Distancia KM':        round(_km, 2),
                        'Chegada':             _cheg.strftime("%H:%M"),
                        'Tempo Descarga (min)': _g.get('descarga', TEMPO_DESCARGA_MIN),
                        'Limite Entrega':      _g['limite'],
                        'Dentro do Prazo':     '✅' if _ok else '❌',
                        'Fonte Rota':          'Cache/Haversine'
                    })

                    _h_ref_db = _cheg + timedelta(minutes=_g.get('descarga', TEMPO_DESCARGA_MIN))
                    _p_ref_db = _dest

                id_trip_db += 1

        df_dbscan = pd.DataFrame(timeline_dbscan)
        if not df_dbscan.empty:
            df_dbscan, co2_dbscan_kg, _co2_db_t = calcular_co2(df_dbscan, consumo_combustivel)
            df_dbscan, frete_dbscan = calcular_frete_total(df_dbscan, tabela_frete)
            salvar_excel(df_dbscan,
                         os.path.join(DRIVE_PATH, '06_TIMELINE_DBSCAN.xlsx'),
                         '06_TIMELINE_DBSCAN.xlsx')
            _nv_db   = df_dbscan['Viagem'].nunique()
            _ok_db   = int((df_dbscan['Dentro do Prazo'] == '✅').sum())
            _oc_db   = df_dbscan['Ocupação %'].mean()
            _km_db   = df_dbscan['Distancia KM'].sum()
            print(f"   ✅ {_nv_db} viagem(ns) DBSCAN → 06_TIMELINE_DBSCAN.xlsx")
            print(f"   👥 Clientes/viagem : {df_dbscan.groupby('Viagem')['Cliente'].count().mean():.1f}")
            print(f"   ⏱️  No prazo        : {_ok_db}/{len(df_dbscan)}")
            print(f"   🚛 Ocupação média  : {_oc_db:.1f}%")
            print(f"   📏 Distância total : {_km_db:.1f} km")
            print(f"   💰 Frete           : R$ {frete_dbscan:,.2f}")
            if co2_dbscan_kg > 0:
                print(f"   🌿 CO2             : {co2_dbscan_kg:.1f} kg")

except ImportError:
    print("   ⚠️  scikit-learn não instalado — DBSCAN ignorado.")
    print("      Execute: pip install scikit-learn")
except Exception as _e_db:
    print(f"   ⚠️  Erro no plano DBSCAN: {repr(_e_db)}")

# =============================================================
# SEÇÃO 13E — PLANO DBSCAN + AG
# K-Means define clusters iniciais → AG reorganiza destinos entre clusters
#
# Estratégia híbrida para TCC:
#   1. Clusters do DBSCAN viram a população inicial do AG (seed inteligente)
#   2. AG é livre para mover destinos entre clusters (crossover OX global)
#   3. Fitness idêntico ao AG puro: frete + penalidade por atraso
#   4. Compara: seed aleatório (AG puro) vs seed geográfico (DBSCAN+AG)
# =============================================================
print("\n🧬📍 Plano DBSCAN + AG — clustering geográfico + otimização evolutiva...")

df_dbscan_ag     = pd.DataFrame()
timeline_dbscan_ag = []
co2_dbscan_ag_kg  = 0.0
frete_dbscan_ag   = 0.0

try:
    from sklearn.cluster import DBSCAN as _DBSCAN2

    _db2_grupos = construir_grupos_destino(eng_final)
    for g in _db2_grupos:
        g['limite'] = limite_otimizado_cliente(g['cli'])
    _db2_n = len(_db2_grupos)

    if _db2_n < 2:
        print("   ⚠️  Poucos destinos — pulando DBSCAN+AG.")
    else:
        # ── 1. Gera clusters DBSCAN (seed geográfico) ────────
        _coords2     = np.array([[g['lat'], g['lon']] for g in _db2_grupos])
        _coords2_rad = np.radians(_coords2)
        _eps2_rad    = RAIO_DBSCAN_KM / 6371.0
        _db2         = _DBSCAN2(eps=_eps2_rad, min_samples=1,
                                algorithm='ball_tree', metric='haversine').fit(_coords2_rad)
        _labels2     = _db2.labels_

        # Monta cromossomo seed: índices ordenados por cluster, depois por urgência
        _cluster_order = {}
        for i, lbl in enumerate(_labels2):
            _cluster_order.setdefault(lbl, []).append(i)

        # Cromossomo seed = destinos ordenados por cluster (vizinhos juntos)
        _seed_cromossomo = []
        for lbl in sorted(_cluster_order.keys()):
            # Dentro de cada cluster, ordena por limite de entrega
            _idxs = sorted(_cluster_order[lbl],
                           key=lambda i: lim_min(_db2_grupos[i]['limite']))
            _seed_cromossomo.extend(_idxs)

        print(f"   📍 Seed geográfico: {len(_cluster_order)} cluster(s) DBSCAN")
        print(f"   🧬 AG com seed inteligente: pop={AG_POPULACAO} | ger={AG_GERACOES}")

        # ── 2. AG com população inicial baseada no seed ───────
        import random as _rnd2
        import copy   as _cp2

        def _ag2_pop_inicial(n_pop, seed):
            """
            População inicial: seed geográfico + variações por mutação leve.
            Garante diversidade mantendo a estrutura geográfica como base.
            """
            pop = [seed[:]]                  # primeiro indivíduo = seed puro
            for _ in range(n_pop - 1):
                ind = seed[:]
                # Aplica 1-3 trocas aleatórias para diversificar
                n_trocas = _rnd2.randint(1, max(1, _db2_n // 4))
                for _ in range(n_trocas):
                    a, b = _rnd2.sample(range(_db2_n), 2)
                    ind[a], ind[b] = ind[b], ind[a]
                pop.append(ind)
            return pop

        # Reutiliza funções do AG puro (já definidas na seção 13C)
        _rnd2.seed(99)   # seed diferente do AG puro para reprodutibilidade independente
        _pop2        = _ag2_pop_inicial(AG_POPULACAO, _seed_cromossomo)
        _fit2_vals   = [_ag_fitness(ind)[0] for ind in _pop2]

        _melhor2_idx = min(range(len(_pop2)), key=lambda i: _fit2_vals[i])
        _melhor2_ind = _pop2[_melhor2_idx][:]
        _melhor2_fit = _fit2_vals[_melhor2_idx]
        _fit2_ini    = _melhor2_fit
        _hist2       = [(0, _melhor2_fit, sum(_fit2_vals)/len(_fit2_vals))]

        print(f"   Geração 0 (seed DBSCAN): melhor fitness = R${_melhor2_fit:,.0f}")

        for _ger2 in range(1, AG_GERACOES + 1):
            _rank2   = sorted(range(len(_pop2)), key=lambda i: _fit2_vals[i])
            _nova2   = [_pop2[i][:] for i in _rank2[:AG_ELITISMO]]

            while len(_nova2) < AG_POPULACAO:
                _p1 = _ag_torneio(_pop2, _fit2_vals)
                if _rnd2.random() < AG_TAXA_CROSSOVER:
                    _p2   = _ag_torneio(_pop2, _fit2_vals)
                    _filho2 = _ag_crossover_ox(_p1, _p2)
                else:
                    _filho2 = _p1[:]
                if _rnd2.random() < AG_TAXA_MUTACAO:
                    _filho2 = _ag_mutar(_filho2)
                _nova2.append(_filho2)

            _pop2      = _nova2
            _fit2_vals = [_ag_fitness(ind)[0] for ind in _pop2]

            _idx2 = min(range(len(_pop2)), key=lambda i: _fit2_vals[i])
            if _fit2_vals[_idx2] < _melhor2_fit:
                _melhor2_fit = _fit2_vals[_idx2]
                _melhor2_ind = _pop2[_idx2][:]

            _med2 = sum(_fit2_vals) / len(_fit2_vals)
            _hist2.append((_ger2, _melhor2_fit, _med2))

            if _ger2 % 20 == 0 or _ger2 == AG_GERACOES:
                print(f"   Geração {_ger2:3d}: melhor = R${_melhor2_fit:,.0f}  "
                      f"(melhoria: {(_fit2_ini-_melhor2_fit)/_fit2_ini*100:+.1f}%)")

        print(f"\n   🏆 Melhor DBSCAN+AG: R${_melhor2_fit:,.0f}  "
              f"(AG puro era R${melhor_fit:,.0f})")

        # ── 3. Gráfico de convergência DBSCAN+AG ─────────────
        _gerar_grafico_convergencia(
            _hist2, _fit2_ini, _melhor2_fit,
            os.path.join(DRIVE_PATH, 'DBSCAN_AG_convergencia.html')
        )

        # ── 4. Decodifica melhor indivíduo em timeline ────────
        _trips2_ag   = _ag_decodificar(_melhor2_ind)
        id_trip_db2  = 1
        _h_db2       = datetime.today().replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
        _prox_doca2  = _h_db2
        _prox_patio2 = _h_db2

        for _trip2 in _trips2_ag:
            _trk2, _aloc2, _sob2 = empacotar_em_caminhao(_trip2, truck_list)
            if not _trk2 or not _aloc2:
                continue

            _qtd2  = sum(len(g['engs']) for g in _trip2)
            _tc2   = tempo_carregamento(_qtd2)
            _loc2  = _prox_carga(_trk2['tipo'])
            if _loc2 == 'Pateo':
                _hi2         = _prox_patio2
                _hs2         = _hi2 + timedelta(minutes=_tc2)
                _prox_patio2 = _hs2
            else:
                _hi2        = _prox_doca2
                _hs2        = _hi2 + timedelta(minutes=_tc2)
                _prox_doca2 = _hs2

            _rot2, _, _cheg2 = simular_viagem(
                _trip2, mem_cache, gmaps, AG_USAR_CACHE_API, controlador_api,
                h_inicio=_hs2)

            _vol2 = min(
                sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                    if 'dim_f' in e else e.get('vol', 0) for e in _aloc2),
                _trk2['vol'])

            _cm2 = {}
            for _e in _aloc2:
                _cm2.setdefault(_e['cli'], []).append(_e)

            _pr2 = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
            _hr2 = _hs2
            _tid2 = f"DBA_{id_trip_db2:03d}"

            for _g in _trip2:
                _cli  = _g['cli']
                _engs = _cm2.get(_cli, [])
                if not _engs:
                    continue
                _dst2    = (_g['lat'], _g['lon'])
                _rid2    = _rota_id(_pr2, _dst2)
                _km2, _m2 = _rot2.get(_rid2, _haversine(_pr2, _dst2))
                _cb2     = _hr2 + timedelta(minutes=_m2)
                _ca2     = ajustar_pausas(_hr2, _cb2)
                _ok2     = dt_min(_ca2) <= lim_min(_g['limite'])

                timeline_dbscan_ag.append({
                    'Viagem':              _tid2,
                    'Caminhão':            _trk2['tipo'],
                    'Vol Caminhão (m³)':   round(_trk2['vol'], 2),
                    'Vol Carga Total (m³)': round(_vol2, 4),
                    'Ocupação %':          round(_vol2 / _trk2['vol'] * 100, 1),
                    'Cliente':             _cli,
                    'Qtd Engradados':      len(_engs),
                    'Tipos Engradados':    ", ".join(sorted(set(e['tipo'] for e in _engs))),
                    'NFs':                 "/".join(sorted(set(e['nf'] for e in _engs))),
                    'Início Carga':        _hi2.strftime("%H:%M"),
                    'Local Carga':         _loc2,
                    'Tempo Carga (min)':   _tc2,
                    'Saída Armazém':       _hs2.strftime("%H:%M"),
                    'Distancia KM':        round(_km2, 2),
                    'Chegada':             _ca2.strftime("%H:%M"),
                    'Tempo Descarga (min)': _g.get('descarga', TEMPO_DESCARGA_MIN),
                    'Limite Entrega':      _g['limite'],
                    'Dentro do Prazo':     '✅' if _ok2 else '❌',
                    'Fonte Rota':          'Cache/Haversine'
                })

                _hr2 = _ca2 + timedelta(minutes=_g.get('descarga', TEMPO_DESCARGA_MIN))
                _pr2 = _dst2

            id_trip_db2 += 1

        df_dbscan_ag = pd.DataFrame(timeline_dbscan_ag)
        if not df_dbscan_ag.empty:
            df_dbscan_ag, co2_dbscan_ag_kg, _co2_dba_t = calcular_co2(
                df_dbscan_ag, consumo_combustivel)
            df_dbscan_ag, frete_dbscan_ag = calcular_frete_total(
                df_dbscan_ag, tabela_frete)
            salvar_excel(df_dbscan_ag,
                         os.path.join(DRIVE_PATH, '06_TIMELINE_DBSCAN_AG.xlsx'),
                         '06_TIMELINE_DBSCAN_AG.xlsx')
            _nv_dba  = df_dbscan_ag['Viagem'].nunique()
            _ok_dba  = int((df_dbscan_ag['Dentro do Prazo'] == '✅').sum())
            _oc_dba  = df_dbscan_ag['Ocupação %'].mean()
            _km_dba  = df_dbscan_ag['Distancia KM'].sum()
            print(f"   ✅ {_nv_dba} viagem(ns) DBSCAN+AG → 06_TIMELINE_DBSCAN_AG.xlsx")
            print(f"   👥 Clientes/viagem : {df_dbscan_ag.groupby('Viagem')['Cliente'].count().mean():.1f}")
            print(f"   ⏱️  No prazo        : {_ok_dba}/{len(df_dbscan_ag)}")
            print(f"   🚛 Ocupação média  : {_oc_dba:.1f}%")
            print(f"   📏 Distância total : {_km_dba:.1f} km")
            print(f"   💰 Frete           : R$ {frete_dbscan_ag:,.2f}")
            if co2_dbscan_ag_kg > 0:
                print(f"   🌿 CO2             : {co2_dbscan_ag_kg:.1f} kg")

except ImportError:
    print("   ⚠️  scikit-learn não instalado — DBSCAN+AG ignorado.")
    print("      Execute: pip install scikit-learn")
except Exception as _e_dba:
    print(f"   ⚠️  Erro no plano DBSCAN+AG: {repr(_e_dba)}")

# =============================================================
# 13B. TABELA DE FRETE E GRAFICO COMPARATIVO
# =============================================================
print("\n💰 Calculando custos de frete...")
df_tl,  frete_real = calcular_frete_total(df_tl,  tabela_frete)
df_otm, frete_otm  = calcular_frete_total(df_otm, tabela_frete)

salvar_excel(df_tl,  os.path.join(DRIVE_PATH, "06_TIMELINE_REAL.xlsx"),       "06_TIMELINE_REAL.xlsx")
salvar_excel(df_otm, os.path.join(DRIVE_PATH, "06_TIMELINE_OTIMIZADO.xlsx"),   "06_TIMELINE_OTIMIZADO.xlsx")

# =============================================================
# PLANO DE CARGA
# =============================================================
print("\n📋 Gerando Plano de Carga...")

def gerar_plano_carga(timeline_data, eng_final_data, caminho_saida, label=''):
    if not timeline_data:
        print(f"   ⚠️  {label}Timeline vazio.")
        return

    eng_por_nf = {}
    for eng in eng_final_data:
        eng_por_nf.setdefault(eng['nf'], []).append(eng)

    viagens_ord = {}
    for r in timeline_data:
        viagens_ord.setdefault(r['Viagem'], []).append(r)

    col_cam = 'Caminhão' if timeline_data and 'Caminhão' in timeline_data[0] else 'Caminhao'

    try:
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            resumo_rows = []

            for trip_id, paradas in sorted(viagens_ord.items()):
                caminhao    = paradas[0].get(col_cam, '')
                linhas      = []
                ordem_carga = 1

                for p in paradas:
                    destinatario = p['Cliente']
                    nfs_parada   = [nf.strip() for nf in str(p['NFs']).split('/')]

                    for nf in nfs_parada:
                        engs_nf   = eng_por_nf.get(nf, [])
                        engs_dest = [e for e in engs_nf
                                     if e['cli'] == destinatario or
                                     destinatario.startswith(e['cli'][:20])]
                        if not engs_dest:
                            engs_dest = engs_nf

                        for eng in sorted(engs_dest, key=lambda e: e['id']):
                            linhas.append({
                                'Trip':             trip_id,
                                'Caminhao':         caminhao,
                                'Ordem Carga':      ordem_carga,
                                'ID Engradado':     eng['id'],
                                'Tipo Engradado':   eng['tipo'],
                                'Nota Fiscal':      nf,
                                'Destinatario':     destinatario,
                                'Chegada Prevista': p.get('Chegada', ''),
                                'Limite Entrega':   p.get('Limite Entrega', ''),
                                'Dentro do Prazo':  p.get('Dentro do Prazo', ''),
                            })
                            ordem_carga += 1

                if not linhas:
                    continue

                df_trip = pd.DataFrame(linhas)
                aba     = trip_id[:31]
                df_trip.to_excel(writer, sheet_name=aba, index=False)

                ws = writer.sheets[aba]
                for col in ws.columns:
                    max_len = max(len(str(col[0].value or '')),
                                  *[len(str(c.value or '')) for c in col[1:]])
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 45)

                resumo_rows.append({
                    'Trip':           trip_id,
                    'Caminhao':       caminhao,
                    'Qtd Engradados': len(linhas),
                    'Destinatarios':  len(paradas),
                    'NFs':            "/".join(sorted(set(r['NFs'] for r in paradas))),
                })

            if resumo_rows:
                df_resumo = pd.DataFrame(resumo_rows)
                df_resumo.to_excel(writer, sheet_name='RESUMO', index=False)
                ws_res = writer.sheets['RESUMO']
                for col in ws_res.columns:
                    max_len = max(len(str(col[0].value or '')),
                                  *[len(str(c.value or '')) for c in col[1:]])
                    ws_res.column_dimensions[col[0].column_letter].width = min(max_len + 3, 50)

        print(f"   ✅ {label}Plano de Carga → {caminho_saida} "
              f"({len(viagens_ord)} abas + RESUMO)")

    except PermissionError:
        print(f"   ❌ Arquivo aberto no Excel — feche e rode novamente.")
    except Exception as e_pc:
        print(f"   ⚠️  Erro ao gerar Plano de Carga: {repr(e_pc)}")


gerar_plano_carga(timeline,     eng_final,
                  os.path.join(DRIVE_PATH, 'Plano de Carga.xlsx'),
                  label='Original — ')

gerar_plano_carga(timeline_otm, eng_final,
                  os.path.join(DRIVE_PATH, 'Plano de Carga Otimizado.xlsx'),
                  label='Otimizado — ')

# Imprime comparativo completo agora que frete está disponível
_imprimir_comparativo(
    fr_real=frete_real,
    fr_otm=frete_otm,
    fr_ag=globals().get('frete_ag') if ('df_ag' in globals() and not df_ag.empty) else None
)

# --- Grafico comparativo ---
print("\n📊 Gerando grafico comparativo...")
try:
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches

    COR_ORIG = "#E63946"
    COR_OTM  = "#2A9D8F"
    fig, axes = plt.subplots(1, 3, figsize=(15, 6))
    fig.suptitle(
        "Impacto da Janela de Entrega na Eficiencia Logistica\n"
        "Original vs Otimizado (ate 02:00 do dia seguinte)",
        fontsize=13, fontweight="bold", y=1.02
    )

    def painel(ax, titulo, ylabel, vals, labels, fmt_fn, nota):
        bars = ax.bar(labels, vals, color=[COR_ORIG, COR_OTM],
                      width=0.5, edgecolor="white", linewidth=1.5)
        ax.set_title(titulo, fontsize=11, fontweight="bold", pad=10)
        ax.set_ylabel(ylabel, fontsize=10)
        mx = max(vals) if max(vals) > 0 else 1
        ax.set_ylim(0, mx * 1.35)
        for bar, val in zip(bars, vals):
            ax.text(bar.get_x() + bar.get_width()/2,
                    bar.get_height() + mx * 0.02,
                    fmt_fn(val), ha="center", va="bottom",
                    fontweight="bold", fontsize=11)
        ax.text(0.5, 0.93, nota, transform=ax.transAxes,
                ha="center", fontsize=9, color="#444", style="italic")
        ax.spines[["top","right"]].set_visible(False)
        ax.yaxis.grid(True, alpha=0.3)
        ax.set_axisbelow(True)

    red_cam = n_caminhoes_real - n_caminhoes_otm
    pct_cam = red_cam / n_caminhoes_real * 100 if n_caminhoes_real > 0 else 0
    painel(axes[0], "Caminhoes utilizados", "Quantidade",
           [n_caminhoes_real, n_caminhoes_otm], ["Original","Otimizado"],
           lambda v: str(int(v)),
           f"Reducao: {red_cam} ({pct_cam:.0f}%)")

    red_co2 = co2_real_kg - co2_otm_kg
    pct_co  = red_co2 / co2_real_kg * 100 if co2_real_kg > 0 else 0
    painel(axes[1], "Emissao de CO2 (kg)", "kg CO2",
           [co2_real_kg, co2_otm_kg], ["Original","Otimizado"],
           lambda v: f"{v:.0f}",
           f"Reducao: {red_co2:.0f} kg ({pct_co:.0f}%)")

    red_fr  = frete_real - frete_otm
    pct_fr2 = red_fr / frete_real * 100 if frete_real > 0 else 0
    painel(axes[2], "Custo de Frete (R$)", "R$",
           [frete_real, frete_otm], ["Original","Otimizado"],
           lambda v: f"R${v:,.0f}",
           f"Reducao: R${red_fr:,.0f} ({pct_fr2:.0f}%)")

    patch_orig = mpatches.Patch(color=COR_ORIG, label="Original (janela cadastrada)")
    patch_otm  = mpatches.Patch(color=COR_OTM,  label="Otimizado (ate 02:00 +1 dia)")
    fig.legend(handles=[patch_orig, patch_otm], loc="lower center",
               ncol=2, fontsize=10, frameon=False, bbox_to_anchor=(0.5, -0.04))
    plt.tight_layout()
    caminho_graf = os.path.join(DRIVE_PATH, "comparativo_original_otimizado.png")
    plt.savefig(caminho_graf, dpi=150, bbox_inches="tight")
    plt.show()
    print(f"   Grafico salvo: {caminho_graf}")
except Exception as e_graf:
    print(f"   Erro ao gerar grafico: {repr(e_graf)}")

# =============================================================
# 14. MAPA DE ROTAS (Folium)
# =============================================================
print("\n🗺️  Gerando mapa de rotas...")

try:
    import folium

    CORES_VIAGEM = [
        '#E63946','#F4A261','#2A9D8F','#457B9D','#6A0572',
        '#F77F00','#118AB2','#06D6A0','#EF476F','#FFD166',
        '#8338EC','#3A86FF','#FB5607','#FFBE0B','#FF006E',
        '#8AC926','#1982C4','#6A4C93','#FF595E','#6BCB77',
        '#4D908E','#F94144','#90BE6D'
    ]

    mapa = folium.Map(
        location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
        zoom_start=9,
        tiles='OpenStreetMap'
    )

    # Armazém
    folium.Marker(
        location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
        popup=folium.Popup('<b>Armazém Suzano</b>', max_width=200),
        tooltip='Armazém Suzano',
        icon=folium.Icon(color='black', icon='home', prefix='fa')
    ).add_to(mapa)

    # Coordenadas por cliente
    coord_cli = {}
    for e in eng_final:
        if e['cli'] not in coord_cli and e['lat'] != 0:
            coord_cli[e['cli']] = (e['lat'], e['lon'])

    # Agrupa paradas por viagem
    viagens_map = {}
    for r in timeline:
        viagens_map.setdefault(r['Viagem'], []).append(r)

    for idx, (trip, paradas) in enumerate(sorted(viagens_map.items())):
        cor = CORES_VIAGEM[idx % len(CORES_VIAGEM)]

        # Linha da rota: armazém → clientes em ordem de entrega
        pontos = [(ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])]
        for p in paradas:
            coord = coord_cli.get(p['Cliente'])
            if coord:
                pontos.append(coord)

        if len(pontos) > 1:
            # Tenta obter rota pelas rodovias via Google Directions API
            rota_real = []
            # Cache de polylines para não chamar Directions API mais de 1x por rota
            chave_dir = str(sorted([str(p) for p in pontos]))
            if chave_dir in _cache_directions:
                rota_real = _cache_directions[chave_dir]
            elif API_DISPONIVEL and not controlador_api.bloqueado:
                try:
                    # Registra consumo: Directions API custa US$0.005 por rota
                    n_elem_dir = max(1, len(pontos) - 1)
                    if controlador_api.pode_chamar(n_elem_dir):
                        dir_res = gmaps.directions(
                            origin=pontos[0],
                            destination=pontos[-1],
                            waypoints=pontos[1:-1] if len(pontos) > 2 else None,
                            mode="driving"
                        )
                        if dir_res:
                            import polyline as polyline_lib
                            for leg in dir_res[0]['legs']:
                                for step in leg['steps']:
                                    pts = polyline_lib.decode(step['polyline']['points'])
                                    rota_real.extend(pts)
                            controlador_api.registrar(1, n_elem_dir, f'directions_{trip}')
                            _cache_directions[chave_dir] = rota_real
                except Exception:
                    pass   # fallback para linha reta

            if rota_real:
                folium.PolyLine(
                    locations=rota_real,
                    color=cor,
                    weight=3,
                    opacity=0.85,
                    tooltip=f"{trip} — {len(paradas)} parada(s)"
                ).add_to(mapa)
            else:
                folium.PolyLine(
                    locations=pontos,
                    color=cor,
                    weight=3,
                    opacity=0.85,
                    tooltip=f"{trip} — {len(paradas)} parada(s) (linha reta)"
                ).add_to(mapa)

        # Marcadores dos clientes
        for p in paradas:
            coord = coord_cli.get(p['Cliente'])
            if not coord:
                continue
            popup_html = (
                f"<b>{p['Cliente']}</b><br>"
                f"Viagem: {trip}<br>"
                f"NFs: {p['NFs']}<br>"
                f"Engradados: {p['Qtd Engradados']}<br>"
                f"Chegada: {p['Chegada']} (limite: {p['Limite Entrega']})<br>"
                f"Distância: {p['Distancia KM']} km"
            )
            folium.CircleMarker(
                location=coord,
                radius=8,
                color=cor,
                fill=True,
                fill_color=cor,
                fill_opacity=0.9,
                popup=folium.Popup(popup_html, max_width=280),
                tooltip=f"{trip} — {p['Cliente'][:30]}"
            ).add_to(mapa)

    # Legenda
    linhas_leg = ''.join(
        f'<span style="color:{CORES_VIAGEM[i % len(CORES_VIAGEM)]}">&#9644;</span> '
        f'{trip} ({len(viagens_map[trip])} parada{"s" if len(viagens_map[trip])>1 else ""})<br>'
        for i, trip in enumerate(sorted(viagens_map.keys()))
    )
    legenda = (
        '<div style="position:fixed;bottom:30px;left:30px;z-index:1000;'
        'background:white;padding:12px;border-radius:8px;border:1px solid #ccc;'
        'font-size:12px;max-height:320px;overflow-y:auto;'
        'box-shadow:2px 2px 6px rgba(0,0,0,.3)">'
        f'<b>Viagens</b><br>{linhas_leg}</div>'
    )
    mapa.get_root().html.add_child(folium.Element(legenda))

    # ── Marcadores interestaduais no mapa Real ────────────────
    if 'eng_interestadual' in globals() and eng_interestadual:
        _inter_fg = folium.FeatureGroup(name='🚛 Interestaduais (rodovia)', show=True)
        _inter_clis_visto = set()
        for _ei in eng_interestadual:
            _cli_ei = _ei.get('cli','')
            if _cli_ei in _inter_clis_visto:
                continue
            _inter_clis_visto.add(_cli_ei)
            _uf_ei  = _ei.get('_uf_classificada', _ei.get('uf','?'))
            _lat_ei = _ei.get('lat', 0)
            _lon_ei = _ei.get('lon', 0)
            if not _lat_ei or not _lon_ei:
                continue
            _popup_ei = (f"<b>🚛 {_cli_ei}</b><br>"
                        f"UF: <b>{_uf_ei}</b><br>"
                        f"NF: {_ei.get('nf','')}<br>"
                        f"Transportadora terceirizada (rodovia)")
            folium.Marker(
                location=[_lat_ei, _lon_ei],
                popup=folium.Popup(_popup_ei, max_width=240),
                tooltip=f"🚛 {_cli_ei} ({_uf_ei}) — interestadual",
                icon=folium.Icon(color='orange', icon='truck', prefix='fa')
            ).add_to(_inter_fg)
            # Linha tracejada laranja — indica rota rodoviária estimada
            # (linha reta = estimativa; rota real definida pela transportadora)
            folium.PolyLine(
                locations=[
                    [ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
                    [_lat_ei, _lon_ei]
                ],
                color='#f97316', weight=1.5,
                dash_array='8 6', opacity=0.5,
                tooltip='Rota estimada — definida pela transportadora'
            ).add_to(_inter_fg)
        _inter_fg.add_to(mapa)
        folium.LayerControl(collapsed=False).add_to(mapa)

    caminho_mapa = os.path.join(DRIVE_PATH, 'mapa_rotas.html')
    mapa.save(caminho_mapa)
    print(f"   ✅ Mapa salvo → {caminho_mapa}")
    salvar_e_abrir(caminho_mapa)
    print("   🌐 Mapa aberto no navegador padrão.")

except ImportError:
    print("   ⚠️  folium não instalado. Execute: pip install folium")
except Exception as e_mapa:
    print(f"   ⚠️  Erro ao gerar mapa: {repr(e_mapa)}")

# =============================================================
# 14B. MAPA OTIMIZADO
# =============================================================
print("\n🗺️  Gerando mapa otimizado (janela 02:00 do dia seguinte)...")

try:
    import folium

    mapa_otm = folium.Map(
        location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
        zoom_start=9, tiles='OpenStreetMap'
    )

    # Título explicativo
    titulo_html = (
        f'<div style="position:fixed;top:10px;left:50%;transform:translateX(-50%);'
        f'z-index:1000;background:white;padding:10px 20px;border-radius:8px;'
        f'border:2px solid #2A9D8F;font-size:13px;font-weight:bold;'
        f'box-shadow:2px 2px 6px rgba(0,0,0,.3);text-align:center">'
        f'Plano Otimizado - Janela ate 02:00 do dia seguinte<br>'
        f'<span style="font-weight:normal;font-size:11px">'
        f'{viagens_otm} viagens vs {viagens} no original '
        f'({viagens - viagens_otm} caminhoes a menos)</span></div>'
    )
    mapa_otm.get_root().html.add_child(folium.Element(titulo_html))

    folium.Marker(
        location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
        popup=folium.Popup('<b>Armazem Suzano</b>', max_width=200),
        tooltip='Armazem Suzano',
        icon=folium.Icon(color='black', icon='home', prefix='fa')
    ).add_to(mapa_otm)

    viagens_map_otm = {}
    for r in timeline_otm:
        viagens_map_otm.setdefault(r['Viagem'], []).append(r)

    for idx, (trip, paradas) in enumerate(sorted(viagens_map_otm.items())):
        cor = CORES_VIAGEM[idx % len(CORES_VIAGEM)]
        pontos = [(ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])]
        for p in paradas:
            coord = coord_cli.get(p['Cliente'])
            if coord:
                pontos.append(coord)
        if len(pontos) > 1:
            rota_real = []
            chave_dir = str(sorted([str(p) for p in pontos]))
            if chave_dir in _cache_directions:
                rota_real = _cache_directions[chave_dir]
            elif API_DISPONIVEL and not controlador_api.bloqueado:
                try:
                    n_elem_dir = max(1, len(pontos) - 1)
                    if controlador_api.pode_chamar(n_elem_dir):
                        dir_res = gmaps.directions(
                            origin=pontos[0], destination=pontos[-1],
                            waypoints=pontos[1:-1] if len(pontos) > 2 else None,
                            mode='driving'
                        )
                        if dir_res:
                            import polyline as polyline_lib
                            for leg in dir_res[0]['legs']:
                                for step in leg['steps']:
                                    pts = polyline_lib.decode(step['polyline']['points'])
                                    rota_real.extend(pts)
                            controlador_api.registrar(1, n_elem_dir, f'dir_otm_{trip}')
                            _cache_directions[chave_dir] = rota_real
                except Exception:
                    pass
            n_cli = len(set(p['Cliente'] for p in paradas))
            folium.PolyLine(
                locations=rota_real if rota_real else pontos,
                color=cor, weight=4, opacity=0.9,
                tooltip=f'{trip} — {n_cli} cliente(s)'
                    + ('' if rota_real else ' (linha reta)')
            ).add_to(mapa_otm)
        for p in paradas:
            coord = coord_cli.get(p['Cliente'])
            if not coord:
                continue
            dia = p.get('Dia Entrega', 'Mesmo dia')
            popup_html = (
                f"<b>{p['Cliente']}</b><br>"
                f"Viagem: {trip}<br>NFs: {p['NFs']}<br>"
                f"Engradados: {p['Qtd Engradados']}<br>"
                f"Chegada: {p['Chegada']} ({dia})<br>"
                f"Limite: {p['Limite Entrega']}<br>"
                f"Distancia: {p['Distancia KM']} km"
            )
            folium.CircleMarker(
                location=coord, radius=10,
                color='#FF6B35' if dia == 'Dia seguinte' else cor,
                weight=2, fill=True, fill_color=cor, fill_opacity=0.85,
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"{trip} — {p['Cliente'][:30]} ({dia})"
            ).add_to(mapa_otm)

    co2_otm_fmt  = f'{co2_otm_kg:.0f} kg'  if co2_otm_kg  > 0 else 'N/D'
    co2_real_fmt = f'{co2_real_kg:.0f} kg' if co2_real_kg > 0 else 'N/D'
    linhas_leg_otm = ''.join(
        f'<span style="color:{CORES_VIAGEM[i % len(CORES_VIAGEM)]}">&#9644;</span> '
        f'{trip} ({len(viagens_map_otm[trip])} parada'
        f'{"s" if len(viagens_map_otm[trip])>1 else ""})<br>'
        for i, trip in enumerate(sorted(viagens_map_otm.keys()))
    )
    legenda_otm = (
        '<div style="position:fixed;bottom:30px;left:30px;z-index:1000;'
        'background:white;padding:12px;border-radius:8px;border:1px solid #2A9D8F;'
        'font-size:12px;max-height:360px;overflow-y:auto;'
        'box-shadow:2px 2px 6px rgba(0,0,0,.3)">'
        f'<b>Plano Otimizado</b><br>'
        f'<span style="color:#666;font-size:11px">'
        f'Original: {viagens} viagens | CO2: {co2_real_fmt}<br>'
        f'Otimizado: {viagens_otm} viagens | CO2: {co2_otm_fmt}</span><br><br>'
        f'{linhas_leg_otm}'
        '<br><span style="font-size:10px;color:#FF6B35">'
        '&#9711; Borda laranja = entrega dia seguinte</span></div>'
    )
    mapa_otm.get_root().html.add_child(folium.Element(legenda_otm))

    caminho_mapa_otm = os.path.join(DRIVE_PATH, 'mapa_rotas_otimizado.html')
    mapa_otm.save(caminho_mapa_otm)
    print(f'   ✅ Mapa otimizado salvo → {caminho_mapa_otm}')
    salvar_e_abrir(caminho_mapa_otm)
    print('   🌐 Mapa otimizado aberto no navegador padrão.')

except ImportError:
    print('   ⚠️  folium nao instalado. Execute: pip install folium')
except Exception as e_mapa_otm:
    print(f'   ⚠️  Erro ao gerar mapa otimizado: {repr(e_mapa_otm)}')

# =============================================================
# 15. VISUALIZAÇÕES 3D — TOP 3 ENGRADADOS + TOP 3 CAMINHÕES (REAL e OTM)
# =============================================================
print("\n🎨 Gerando visualizações 3D...")

def reconstruir_lay_engradado(rec):
    tipo_eng  = next((e for e in TIPOS_ENGRADADOS
                      if e['nome'] == rec['Tipo Engradado']), TIPOS_ENGRADADOS[0])
    qtd_total = rec['Qtd Caixas Engradado']
    itens_viz = []
    for pn_raw in rec['Part Numbers'].split(', '):
        info = indice_pn.get(normalizar_pn(pn_raw.strip()))
        if info:
            dc, dl, da, ql = info
            itens_viz.append({'dim': [dc, dl, da],
                               'vol': round(dc*dl*da, 6),
                               'pn': pn_raw.strip(), 'qtd_lote': ql})
    base = list(itens_viz)
    while len(itens_viz) < qtd_total and base:
        itens_viz.append(dict(base[len(itens_viz) % len(base)]))
    lay, _ = motor_tetris(tipo_eng['dim'], itens_viz[:qtd_total], passo_x=0.2)
    return lay, tipo_eng

def reconstruir_lay_caminhao(trip_id, tl_data, eng_data):
    nfs_trip = set()
    for r in tl_data:
        if r['Viagem'] == trip_id:
            for nf_t in str(r['NFs']).split('/'):
                nfs_trip.add(nf_t.strip())
    engs_trip = [e for e in eng_data if e['nf'] in nfs_trip]
    if not engs_trip:
        return None, None, []
    col_cam   = 'Caminhão' if tl_data and 'Caminhão' in tl_data[0] else 'Caminhao'
    tipo_nome = next((r[col_cam] for r in tl_data if r['Viagem'] == trip_id), None)
    truck_viz = next((t for t in truck_list if t['tipo'] == tipo_nome), truck_list[-1])
    pecas     = [{**e} for e in engs_trip]
    passo_t   = max(0.1, min(p['dim'][0] for p in pecas))
    sem_e_viz = not permite_empilhamento(tipo_nome or '')
    lay, _    = motor_tetris(truck_viz['dim'], pecas, passo_x=passo_t,
                              sem_empilhamento=sem_e_viz)
    return truck_viz, tipo_nome, lay


# =============================================================
# VISUALIZAÇÃO 3D EM HTML (Three.js) — engradados e caminhões
# =============================================================
def _cor_hex_3d(idx, total, sat=0.72, val=0.82):
    """Gera cor hex HSV para paleta de clientes/PNs."""
    import colorsys
    r, g, b = colorsys.hsv_to_rgb(idx / max(total, 1), sat, val)
    return '0x{:02x}{:02x}{:02x}'.format(int(r*255), int(g*255), int(b*255))


def _html_3d_base(titulo, subtitulo, cubos, cam_dim=None):
    """
    Gera HTML completo com visualização 3D interativa (Three.js).
    cubos: lista de dict {x, y, z, dx, dy, dz, cor_hex, label, wireframe}
    cam_dim: [C, L, A] do caminhão para desenhar o contorno do baú (opcional)
    """
    import json
    cubos_json  = json.dumps(cubos)
    cam_json    = json.dumps(cam_dim) if cam_dim else 'null'
    hoje        = __import__('datetime').datetime.now().strftime('%d/%m/%Y %H:%M')

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>{titulo}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#1a1a2e;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     overflow:hidden;color:white}}
#info{{position:fixed;top:0;left:0;right:0;background:rgba(0,0,0,0.55);
      backdrop-filter:blur(8px);padding:10px 20px;z-index:10;
      display:flex;align-items:center;justify-content:space-between}}
#info h1{{font-size:15px;font-weight:600;color:#e0f7f0}}
#info p{{font-size:11px;color:#aaa;margin-top:2px}}
#info-right{{font-size:11px;color:#888;text-align:right}}
#hint{{position:fixed;bottom:16px;left:50%;transform:translateX(-50%);
      background:rgba(0,0,0,0.6);padding:6px 16px;border-radius:20px;
      font-size:11px;color:#aaa;pointer-events:none;z-index:10}}
#legend{{position:fixed;right:16px;top:60px;background:rgba(0,0,0,0.55);
         backdrop-filter:blur(6px);border-radius:10px;padding:10px 14px;
         max-height:70vh;overflow-y:auto;z-index:10;min-width:180px}}
#legend h2{{font-size:11px;font-weight:600;color:#aaa;text-transform:uppercase;
            letter-spacing:.06em;margin-bottom:8px}}
.leg-item{{display:flex;align-items:center;gap:7px;margin-bottom:6px;font-size:11px}}
.leg-swatch{{width:14px;height:14px;border-radius:3px;flex-shrink:0;
             border:1px solid rgba(255,255,255,0.2)}}
#tooltip{{position:fixed;background:rgba(15,15,15,0.90);color:#eee;
          border-radius:8px;padding:8px 12px;font-size:12px;line-height:1.7;
          pointer-events:none;z-index:20;display:none;max-width:240px;
          box-shadow:0 4px 20px rgba(0,0,0,0.5)}}
canvas{{display:block}}
</style>
</head>
<body>
<div id="info">
  <div>
    <h1>{titulo}</h1>
    <p>{subtitulo}</p>
  </div>
  <div id="info-right">TTB Logistics · {hoje}<br>Arraste para girar · Scroll para zoom</div>
</div>
<div id="legend"><h2>Legenda</h2><div id="leg-items"></div></div>
<div id="hint">🖱 Arrastar: girar &nbsp;|&nbsp; Scroll: zoom &nbsp;|&nbsp; Botão direito: mover</div>
<div id="tooltip"></div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
const CUBOS   = {cubos_json};
const CAM_DIM = {cam_json};

// Cena
const scene    = new THREE.Scene();
scene.background = new THREE.Color(0x1a1a2e);
scene.fog        = new THREE.Fog(0x1a1a2e, 30, 80);

const W = window.innerWidth, H = window.innerHeight;
const camera = new THREE.PerspectiveCamera(45, W/H, 0.01, 200);

const renderer = new THREE.WebGLRenderer({{antialias:true}});
renderer.setSize(W, H);
renderer.shadowMap.enabled = true;
renderer.setPixelRatio(window.devicePixelRatio);
document.body.appendChild(renderer.domElement);

// Luzes
scene.add(new THREE.AmbientLight(0xffffff, 0.55));
const dl = new THREE.DirectionalLight(0xffffff, 0.9);
dl.position.set(10, 20, 10);
dl.castShadow = true;
scene.add(dl);
scene.add(new THREE.DirectionalLight(0x8899ff, 0.3).translateX(-10).translateY(5));

// Grade no chão
const gridHelper = new THREE.GridHelper(40, 40, 0x333355, 0x222244);
scene.add(gridHelper);

// Contorno do baú (se fornecido)
if (CAM_DIM) {{
  const [C, L, A] = CAM_DIM;
  const geo = new THREE.EdgesGeometry(new THREE.BoxGeometry(C, A, L));
  const mat = new THREE.LineBasicMaterial({{color:0x4488cc, linewidth:2}});
  const box = new THREE.LineSegments(geo, mat);
  box.position.set(C/2, A/2, L/2);
  scene.add(box);
}}

// Cria cubos
const meshes = [];
const legendMap = {{}};
CUBOS.forEach((c, i) => {{
  if (c.wireframe) {{
    // Aresta do container
    const geo = new THREE.EdgesGeometry(new THREE.BoxGeometry(c.dx, c.dz, c.dy));
    const mat = new THREE.LineBasicMaterial({{color: parseInt(c.cor), opacity:0.35, transparent:true}});
    const ls  = new THREE.LineSegments(geo, mat);
    ls.position.set(c.x + c.dx/2, c.z + c.dz/2, c.y + c.dy/2);
    scene.add(ls);
    return;
  }}
  const geo = new THREE.BoxGeometry(c.dx, c.dz, c.dy);
  const mat = new THREE.MeshLambertMaterial({{
    color: parseInt(c.cor), transparent:true, opacity:0.88
  }});
  const mesh = new THREE.Mesh(geo, mat);
  mesh.position.set(c.x + c.dx/2, c.z + c.dz/2, c.y + c.dy/2);
  mesh.castShadow = true;
  mesh.receiveShadow = true;
  mesh.userData = {{label: c.label, cor: c.cor}};
  scene.add(mesh);
  meshes.push(mesh);

  // Arestas do cubo
  const edgeGeo = new THREE.EdgesGeometry(geo);
  const edgeMat = new THREE.LineBasicMaterial({{color:0x000000, opacity:0.25, transparent:true}});
  const edges   = new THREE.LineSegments(edgeGeo, edgeMat);
  mesh.add(edges);

  if (!legendMap[c.label]) legendMap[c.label] = c.cor;
}});

// Legenda
const legEl = document.getElementById('leg-items');
Object.entries(legendMap).forEach(([lbl, cor]) => {{
  const hexStr = '#' + parseInt(cor).toString(16).padStart(6,'0');
  const div = document.createElement('div');
  div.className = 'leg-item';
  div.innerHTML = `<div class="leg-swatch" style="background:${{hexStr}}"></div>
                   <span>${{lbl.length > 22 ? lbl.substring(0,21)+'…' : lbl}}</span>`;
  legEl.appendChild(div);
}});

// Posiciona câmera
const allX = CUBOS.filter(c=>!c.wireframe).map(c=>c.x+c.dx);
const allY = CUBOS.filter(c=>!c.wireframe).map(c=>c.y+c.dy);
const allZ = CUBOS.filter(c=>!c.wireframe).map(c=>c.z+c.dz);
const cx   = (Math.max(...allX.concat(0))) / 2;
const cy   = (Math.max(...allY.concat(0))) / 2;
const cz   = (Math.max(...allZ.concat(0))) / 2;
const dist = Math.max(...allX.concat(0), ...allY.concat(0), ...allZ.concat(0)) * 1.8;
camera.position.set(cx + dist*0.7, cz + dist*0.6, cy + dist*0.9);
camera.lookAt(cx, cz/2, cy);

// Orbit controls manual
let isDragging=false, isRight=false;
let prevX=0, prevY=0;
let theta=0.6, phi=0.9, radius=dist;
let panX=cx, panY=cz/2, panZ=cy;

function updateCamera() {{
  camera.position.set(
    panX + radius * Math.sin(phi) * Math.sin(theta),
    panY + radius * Math.cos(phi),
    panZ + radius * Math.sin(phi) * Math.cos(theta)
  );
  camera.lookAt(panX, panY, panZ);
}}
updateCamera();

renderer.domElement.addEventListener('mousedown', e => {{
  isDragging=true; isRight=e.button===2;
  prevX=e.clientX; prevY=e.clientY;
}});
renderer.domElement.addEventListener('contextmenu', e=>e.preventDefault());
window.addEventListener('mouseup', () => isDragging=false);
window.addEventListener('mousemove', e => {{
  if (!isDragging) return;
  const dx = e.clientX - prevX, dy = e.clientY - prevY;
  prevX=e.clientX; prevY=e.clientY;
  if (isRight) {{
    panX -= dx * 0.01; panY += dy * 0.01;
  }} else {{
    theta -= dx * 0.008;
    phi    = Math.max(0.1, Math.min(Math.PI-0.1, phi + dy*0.008));
  }}
  updateCamera();
}});
renderer.domElement.addEventListener('wheel', e => {{
  radius = Math.max(0.5, radius + e.deltaY * 0.02);
  updateCamera();
}});

// Tooltip via raycasting
const raycaster = new THREE.Raycaster();
const mouse     = new THREE.Vector2();
const tip       = document.getElementById('tooltip');
renderer.domElement.addEventListener('mousemove', e => {{
  mouse.x =  (e.clientX / W) * 2 - 1;
  mouse.y = -(e.clientY / H) * 2 + 1;
  raycaster.setFromCamera(mouse, camera);
  const hits = raycaster.intersectObjects(meshes);
  if (hits.length > 0) {{
    const lbl = hits[0].object.userData.label || '';
    tip.style.display = 'block';
    tip.style.left    = (e.clientX + 14) + 'px';
    tip.style.top     = (e.clientY - 10) + 'px';
    tip.textContent   = lbl;
  }} else {{
    tip.style.display = 'none';
  }}
}});

// Resize
window.addEventListener('resize', () => {{
  camera.aspect = window.innerWidth / window.innerHeight;
  camera.updateProjectionMatrix();
  renderer.setSize(window.innerWidth, window.innerHeight);
}});

// Animação
function animate() {{
  requestAnimationFrame(animate);
  renderer.render(scene, camera);
}}
animate();
</script>
</body>
</html>"""


def gerar_html_3d_engradado(eng_id, lay, escolha, caminho_html):
    """Gera HTML 3D interativo do engradado."""
    import colorsys
    C, L, A = escolha['dim']
    pns_unicos = sorted(set(it.get('pn', 'item') for it in lay))
    n = max(len(pns_unicos), 1)
    cor_map = {pn: _cor_hex_3d(i, n) for i, pn in enumerate(pns_unicos)}

    cubos = []
    # Contorno do engradado
    cubos.append({'x':0,'y':0,'z':0,'dx':C,'dy':L,'dz':A,
                  'cor':'0x4488cc','label':'Engradado','wireframe':True})
    for it in lay:
        if 'pos' not in it or 'dim_f' not in it:
            continue
        px, py, pz = it['pos']
        dc, dl, da = it['dim_f']
        pn = it.get('pn', 'item')
        cubos.append({'x':px,'y':py,'z':pz,'dx':dc,'dy':dl,'dz':da,
                      'cor': cor_map.get(pn, '0x888888'),
                      'label': pn, 'wireframe': False})

    titulo    = f"Engradado {eng_id} — {escolha['nome']}"
    subtitulo = (f"Dimensões: {C:.2f} × {L:.2f} × {A:.2f} m  |  "
                 f"{len(lay)} caixa(s)  |  Vol: {escolha['vol']:.3f} m³")
    html = _html_3d_base(titulo, subtitulo, cubos)
    with open(caminho_html, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"   🌐 Engradado 3D HTML → {__import__('os').path.basename(caminho_html)}")


def gerar_html_3d_caminhao(trip_id, truck, alocados, caminho_html):
    """Gera HTML 3D interativo do caminhão com engradados por cliente."""
    C, L, A = truck['dim']
    clis_unicos = sorted(set(e['cli'] for e in alocados))
    n = max(len(clis_unicos), 1)
    cor_map = {cli: _cor_hex_3d(i, n, sat=0.70, val=0.85)
               for i, cli in enumerate(clis_unicos)}

    cubos = []
    # Contorno do baú
    cubos.append({'x':0,'y':0,'z':0,'dx':C,'dy':L,'dz':A,
                  'cor':'0x4488cc','label':'Baú','wireframe':True})
    for eng in alocados:
        if 'pos' not in eng or 'dim_f' not in eng:
            continue
        px, py, pz = eng['pos']
        dc, dl, da = eng['dim_f']
        cli  = eng['cli'].split('[')[0].strip()
        tipo = eng.get('tipo', '')
        label = f"{cli} ({tipo})"
        cubos.append({'x':px,'y':py,'z':pz,'dx':dc,'dy':dl,'dz':da,
                      'cor': cor_map.get(eng['cli'], '0x888888'),
                      'label': label, 'wireframe': False})

    vol_carga = min(sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                        if 'dim_f' in e else e.get('vol',0)
                        for e in alocados), truck['vol'])
    ocup = round(vol_carga / truck['vol'] * 100, 1)
    n_engs = len(alocados)
    n_half = sum(1 for e in alocados if e.get('tipo') == 'Half')

    titulo    = f"Viagem {trip_id} — {truck['tipo']}"
    subtitulo = (f"{C:.1f} × {L:.1f} × {A:.1f} m  |  "
                 f"{n_engs} engradados ({n_half} Half / {n_engs-n_half} Full)  |  "
                 f"Ocupação: {ocup}%")
    html = _html_3d_base(titulo, subtitulo, cubos, cam_dim=[C, L, A])
    with open(caminho_html, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"   🌐 Caminhão 3D HTML → {__import__('os').path.basename(caminho_html)}")

# ── TOP 3 ENGRADADOS por eficiência ───────────────────────────
top3_engs = sorted(rel_eng,
                   key=lambda r: r['Eficiência Volumétrica %'],
                   reverse=True)[:3]

for rank, rec in enumerate(top3_engs, 1):
    try:
        lay_viz, tipo_eng = reconstruir_lay_engradado(rec)
        eng_id = rec['ID Sequencia']
        ef     = rec['Eficiência Volumétrica %']
        caminho_png  = os.path.join(DRIVE_PATH, f'3D_ENG_top{rank}_{eng_id}.png')
        caminho_html = os.path.join(DRIVE_PATH, f'3D_ENG_top{rank}_{eng_id}.html')
        gerar_imagem_engradado(eng_id, lay_viz, tipo_eng, caminho_png)
        gerar_html_3d_engradado(eng_id, lay_viz, tipo_eng, caminho_html)
        print(f"   📦 Engradado #{rank}: {eng_id} | {rec['Tipo Engradado']} | Efic: {ef}%")
    except Exception as e_e:
        print(f"   ⚠️  Erro engradado #{rank}: {repr(e_e)}")

def gerar_grade_caminhoes(label, tl_df, tl_data, eng_data, prefixo):
    """
    Gera um PNG com TODOS os caminhões do plano em grade de subplots.
    Também salva PNG individual por trip.
    Ordenados por ocupação decrescente.
    """
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D          # noqa: F401
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    from matplotlib.patches import Patch

    if tl_df.empty or 'Ocupação %' not in tl_df.columns:
        print(f"   ⚠️  Timeline {label} vazio.")
        return

    ocup_por_trip = (tl_df.groupby('Viagem')['Ocupação %']
                     .first().sort_values(ascending=False))
    trips_ord = list(ocup_por_trip.index)
    n_trips   = len(trips_ord)
    n_cols    = min(4, n_trips)
    n_rows    = (n_trips + n_cols - 1) // n_cols

    fig = plt.figure(figsize=(5.5 * n_cols, 4.5 * n_rows))
    fig.patch.set_facecolor('#f0f0f0')
    fig.suptitle(f'Plano {label} — Todos os caminhões ({n_trips} viagens)',
                 fontsize=13, fontweight='bold', y=1.01)

    print(f"\n   🚛 {label} — gerando {n_trips} caminhão(s):")

    for idx, trip_id in enumerate(trips_ord, 1):
        try:
            ocup = ocup_por_trip[trip_id]
            truck_viz, tipo_nome, lay_truck = reconstruir_lay_caminhao(
                trip_id, tl_data, eng_data)
            if not lay_truck:
                print(f"      ⚠️  {trip_id}: sem engradados.")
                continue

            C, L, A = truck_viz['dim']
            clientes_unicos = sorted(set(e['cli'] for e in lay_truck))
            cores = _paleta_pns(clientes_unicos)

            # ── Subplot na grade ─────────────────────────────
            ax = fig.add_subplot(n_rows, n_cols, idx, projection='3d')
            ax.set_facecolor('#f8f8f8')
            _cubo_wireframe(ax, 0, 0, 0, C, L, A,
                            cor=(0.55, 0.75, 0.95), alpha=0.07, lw=1.0)
            piso = [[(0,0,0),(C,0,0),(C,L,0),(0,L,0)]]
            ax.add_collection3d(Poly3DCollection(piso,
                                alpha=0.20, facecolors='#ccc',
                                edgecolors='#888', linewidths=0.4))
            for eng in lay_truck:
                if 'pos' not in eng or 'dim_f' not in eng:
                    continue
                px, py, pz = eng['pos']
                dc, dl, da = eng['dim_f']
                _cubo_wireframe(ax, px, py, pz, dc, dl, da,
                                cor=cores[eng['cli']], alpha=0.88)
            ax.set_xlim(0, C); ax.set_ylim(0, L); ax.set_zlim(0, A + 0.05)
            ax.set_xlabel('C(m)', fontsize=6, labelpad=2)
            ax.set_ylabel('L(m)', fontsize=6, labelpad=2)
            ax.set_zlabel('A(m)', fontsize=6, labelpad=2)
            ax.tick_params(labelsize=5)
            ax.view_init(elev=20, azim=-50)
            ax.xaxis.pane.fill = False
            ax.yaxis.pane.fill = False
            ax.zaxis.pane.fill = False
            ax.grid(True, alpha=0.10, linewidth=0.3)
            handles = [Patch(facecolor=cores[c], edgecolor='#555',
                             label=c.split('[')[0].strip()[:22] +
                                   ('…' if len(c.split('[')[0].strip()) > 22 else ''))
                       for c in clientes_unicos]
            ax.legend(handles=handles, fontsize=4.5, title='Clientes',
                      title_fontsize=5, loc='upper left',
                      bbox_to_anchor=(0, 1.02), framealpha=0.80)
            n_engs = len(lay_truck)
            ax.set_title(f"{trip_id} | {tipo_nome}\n"
                         f"{n_engs} eng. | {ocup:.0f}% ocup.",
                         fontsize=7, pad=4)

            # ── PNG individual ───────────────────────────────
            caminho_ind = os.path.join(DRIVE_PATH, f'3D_{prefixo}_{trip_id}.png')
            fig_i = plt.figure(figsize=(11, 6))
            ax_i  = fig_i.add_subplot(111, projection='3d')
            ax_i.set_facecolor('#f4f4f4')
            fig_i.patch.set_facecolor('#f4f4f4')
            _cubo_wireframe(ax_i, 0, 0, 0, C, L, A,
                            cor=(0.55, 0.75, 0.95), alpha=0.07, lw=1.4)
            piso2 = [[(0,0,0),(C,0,0),(C,L,0),(0,L,0)]]
            ax_i.add_collection3d(Poly3DCollection(piso2,
                                  alpha=0.25, facecolors='#ccc',
                                  edgecolors='#888', linewidths=0.5))
            for eng in lay_truck:
                if 'pos' not in eng or 'dim_f' not in eng:
                    continue
                px, py, pz = eng['pos']
                dc, dl, da = eng['dim_f']
                _cubo_wireframe(ax_i, px, py, pz, dc, dl, da,
                                cor=cores[eng['cli']], alpha=0.88)
            ax_i.set_xlim(0, C); ax_i.set_ylim(0, L); ax_i.set_zlim(0, A + 0.1)
            ax_i.set_xlabel('Comprimento (m)', labelpad=6, fontsize=8)
            ax_i.set_ylabel('Largura (m)',     labelpad=6, fontsize=8)
            ax_i.set_zlabel('Altura (m)',      labelpad=6, fontsize=8)
            ax_i.view_init(elev=20, azim=-50)
            ax_i.xaxis.pane.fill = False
            ax_i.yaxis.pane.fill = False
            ax_i.zaxis.pane.fill = False
            ax_i.grid(True, alpha=0.12, linewidth=0.4)
            handles_i = [Patch(facecolor=cores[c], edgecolor='#444',
                               label=c.split('[')[0].strip()[:32] +
                                     ('…' if len(c.split('[')[0].strip()) > 32 else ''))
                         for c in clientes_unicos]
            ax_i.legend(handles=handles_i, loc='upper left', fontsize=7,
                        title='Clientes', title_fontsize=8,
                        bbox_to_anchor=(0.0, 1.0), framealpha=0.88)
            n_half = sum(1 for e in lay_truck if e.get('tipo') == 'Half')
            ax_i.set_title(
                f"Viagem {trip_id}  |  {tipo_nome}"
                f"  ({C:.1f} × {L:.1f} × {A:.1f} m)\n"
                f"{n_engs} engradados ({n_half} Half / {n_engs-n_half} Full)"
                f"  |  Ocupação: {ocup:.1f}%",
                fontsize=10, pad=14)
            fig_i.tight_layout()
            fig_i.savefig(caminho_ind, dpi=130, bbox_inches='tight',
                          facecolor=fig_i.get_facecolor())
            plt.close(fig_i)

            # HTML 3D interativo do caminhão
            caminho_html_ind = os.path.join(DRIVE_PATH,
                                            f'3D_{prefixo}_{trip_id}.html')
            gerar_html_3d_caminhao(trip_id, truck_viz, lay_truck, caminho_html_ind)
            print(f"      ✅ {trip_id} | {tipo_nome} | {ocup:.1f}%")

        except Exception as e_g:
            print(f"      ⚠️  Erro {trip_id}: {repr(e_g)}")

    fig.tight_layout(rect=[0, 0, 1, 0.98])
    caminho_grade = os.path.join(DRIVE_PATH, f'3D_{prefixo}_GRADE.png')
    fig.savefig(caminho_grade, dpi=130, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    plt.show()
    plt.close(fig)
    print(f"   🖼️  Grade salva → {os.path.basename(caminho_grade)}")


gerar_grade_caminhoes('REAL',      df_tl,  timeline,     eng_final, 'REAL')
gerar_grade_caminhoes('OTIMIZADO', df_otm, timeline_otm, eng_final, 'OTM')
# Gera grade 3D do AG (se df_ag existir)
if 'df_ag' in globals() and isinstance(df_ag, pd.DataFrame) and not df_ag.empty:
    gerar_grade_caminhoes('AG', df_ag, timeline_ag, eng_final, 'AG')
else:
    print("   ⚠️  AG: df_ag não disponível — grade 3D AG ignorada.")

# Abre HTMLs dos top3 engradados no navegador
print("\n🌐 Abrindo visualizações 3D HTML no navegador...")
for rank, rec in enumerate(top3_engs, 1):
    eng_id = rec['ID Sequencia']
    h = os.path.join(DRIVE_PATH, f'3D_ENG_top{rank}_{eng_id}.html')
    if os.path.exists(h):
        salvar_e_abrir(h, silent=True)

# Gera página índice HTML com links para todos os caminhões 3D
def gerar_indice_3d(label, df_tl_ref, prefixo, caminho_idx):
    import colorsys, json
    if df_tl_ref.empty:
        return
    col_cam = 'Caminhão' if 'Caminhão' in df_tl_ref.columns else 'Caminhao'
    ocup_col = 'Ocupação %' if 'Ocupação %' in df_tl_ref.columns else None
    trips = sorted(df_tl_ref['Viagem'].unique().tolist())
    hoje  = __import__('datetime').datetime.now().strftime('%d/%m/%Y %H:%M')

    cards = []
    for i, trip in enumerate(trips):
        row0   = df_tl_ref[df_tl_ref['Viagem'] == trip].iloc[0]
        cam    = str(row0[col_cam]) if col_cam in df_tl_ref.columns else ''
        ocup   = f"{float(row0[ocup_col]):.0f}%" if ocup_col else ''
        n_cli  = len(df_tl_ref[df_tl_ref['Viagem'] == trip])
        fname  = f'3D_{prefixo}_{trip}.html'
        r, g, b = colorsys.hsv_to_rgb(i / max(len(trips),1), 0.65, 0.75)
        cor = '#{:02x}{:02x}{:02x}'.format(int(r*255),int(g*255),int(b*255))
        cards.append(f'''<a href="{fname}" class="card" style="border-color:{cor}">
          <div class="card-badge" style="background:{cor}">{trip}</div>
          <div class="card-cam">{cam}</div>
          <div class="card-info">{n_cli} parada(s) · {ocup}</div>
        </a>''')

    html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>3D Índice — {label}</title>
<style>
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     background:#111827;color:#e5e7eb;margin:0;padding:0}}
header{{background:#1D9E75;padding:18px 28px;
        display:flex;align-items:center;justify-content:space-between}}
header h1{{font-size:18px;font-weight:600;color:white}}
header span{{font-size:12px;color:rgba(255,255,255,0.8)}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));
       gap:14px;padding:24px}}
.card{{background:#1f2937;border:2px solid #374151;border-radius:14px;
       padding:16px;text-decoration:none;color:inherit;
       transition:transform .15s,box-shadow .15s;display:block}}
.card:hover{{transform:translateY(-3px);box-shadow:0 8px 24px rgba(0,0,0,0.4)}}
.card-badge{{display:inline-block;padding:3px 10px;border-radius:20px;
             font-size:11px;font-weight:700;color:white;margin-bottom:8px}}
.card-cam{{font-size:15px;font-weight:600;margin-bottom:4px}}
.card-info{{font-size:12px;color:#9ca3af}}
footer{{text-align:center;padding:16px;font-size:11px;color:#4b5563}}
</style></head><body>
<header>
  <h1>🚚 3D Caminhões — {label}</h1>
  <span>TTB Logistics · {hoje} · {len(trips)} viagens</span>
</header>
<div class="grid">{''.join(cards)}</div>
<footer>TTB Logistics AI Simulator · Clique em uma viagem para ver o 3D interativo</footer>
</body></html>"""

    with open(caminho_idx, 'w', encoding='utf-8') as f:
        f.write(html)
    salvar_e_abrir(caminho_idx, silent=True)
    print(f"   🌐 Índice 3D aberto → {os.path.basename(caminho_idx)}")


gerar_indice_3d('Plano Real',      df_tl,  'REAL',
                os.path.join(DRIVE_PATH, '3D_REAL_indice.html'))
gerar_indice_3d('Plano Otimizado', df_otm, 'OTM',
                os.path.join(DRIVE_PATH, '3D_OTM_indice.html'))
if 'df_ag' in globals() and isinstance(df_ag, pd.DataFrame) and not df_ag.empty:
    gerar_indice_3d('Plano AG', df_ag, 'AG',
                    os.path.join(DRIVE_PATH, '3D_AG_indice.html'))
else:
    print("   ⚠️  AG: df_ag não disponível — índice 3D AG ignorado.")

# =============================================================
# TIMELINE VISUAL — diagrama de rotas em linha reta por trip
# =============================================================
def gerar_timeline_visual_html(df_tl, label, caminho_saida):
    """
    Gera arquivo HTML interativo com o timeline visual de todas as trips.
    Abre automaticamente no navegador padrão ao finalizar.
    Cada trip ocupa uma linha horizontal; paradas são marcadas com tooltip.
    """
    import colorsys, webbrowser, json

    if df_tl.empty:
        print(f"   ⚠️  Timeline {label} vazio — visualização não gerada.")
        return

    col_sai = 'Saída Armazém' if 'Saída Armazém' in df_tl.columns else None
    col_cam = 'Caminhão'      if 'Caminhão'      in df_tl.columns else 'Caminhao'
    col_lc  = 'Local Carga'   if 'Local Carga'   in df_tl.columns else None

    # Ordena trips por horário de saída
    if col_sai:
        ordem_trips = (df_tl.groupby('Viagem')[col_sai]
                       .first().sort_values().index.tolist())
    else:
        ordem_trips = sorted(df_tl['Viagem'].unique().tolist())

    n_trips = len(ordem_trips)

    def _hex(i):
        r, g, b = colorsys.hsv_to_rgb(i / max(n_trips, 1), 0.70, 0.78)
        return '#{:02x}{:02x}{:02x}'.format(int(r*255), int(g*255), int(b*255))

    def _hex_d(i):
        r, g, b = colorsys.hsv_to_rgb(i / max(n_trips, 1), 0.80, 0.50)
        return '#{:02x}{:02x}{:02x}'.format(int(r*255), int(g*255), int(b*255))

    # Monta estrutura de dados por trip
    trips_dados = {}
    km_max_global = 0.0
    for trip in ordem_trips:
        df_t    = df_tl[df_tl['Viagem'] == trip].copy()
        row0    = df_t.iloc[0]
        cam     = str(row0[col_cam]) if col_cam in df_t.columns else ''
        saida   = str(row0[col_sai]) if col_sai else ''
        local_c = str(row0[col_lc])  if col_lc  else ''
        km_acum = 0.0
        paradas = []
        for _, row in df_t.iterrows():
            km_seg  = float(row['Distancia KM']) if pd.notna(row['Distancia KM']) else 0
            km_acum = round(km_acum + km_seg, 1)
            cli     = str(row['Cliente']).split('[')[0].strip()
            chegada = str(row['Chegada'])
            limite  = str(row['Limite Entrega']).replace(':00','').strip()
            prazo   = '✅' in str(row['Dentro do Prazo'])
            desc_min= int(row['Tempo Descarga (min)']) if 'Tempo Descarga (min)' in df_t.columns and pd.notna(row.get('Tempo Descarga (min)')) else 0
            nfs     = str(row['NFs']) if 'NFs' in df_t.columns else ''
            paradas.append({'cli': cli, 'chegada': chegada, 'limite': limite,
                            'prazo': prazo, 'km': km_acum, 'nfs': nfs,
                            'descarga': desc_min})
        km_tot = paradas[-1]['km'] if paradas else 0
        km_max_global = max(km_max_global, km_tot)
        ini_carga = str(row0.get('Início Carga', '')) if 'Início Carga' in df_t.columns else ''
        t_carga   = str(row0.get('Tempo Carga (min)', '')) if 'Tempo Carga (min)' in df_t.columns else ''
        ocup      = str(row0.get('Ocupação %', '')) if 'Ocupação %' in df_t.columns else ''
        trips_dados[trip] = {
            'cam': cam, 'saida': saida, 'local': local_c,
            'ini_carga': ini_carga, 't_carga': t_carga, 'ocup': ocup,
            'km_tot': km_tot, 'paradas': paradas
        }

    # Serializa para JSON (injetado no JS)
    dados_json = json.dumps(trips_dados, ensure_ascii=False)
    cores_json = json.dumps({trip: _hex(i) for i, trip in enumerate(ordem_trips)},
                            ensure_ascii=False)
    cores_d_json = json.dumps({trip: _hex_d(i) for i, trip in enumerate(ordem_trips)},
                              ensure_ascii=False)
    ordem_json = json.dumps(ordem_trips, ensure_ascii=False)
    hoje_str   = __import__('datetime').datetime.now().strftime('%d/%m/%Y %H:%M')

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Timeline Visual — {label}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     background:#f2f2f2;color:#1a1a1a}}
header{{background:#1D9E75;color:white;padding:16px 28px;
        display:flex;align-items:center;justify-content:space-between;
        box-shadow:0 2px 8px rgba(0,0,0,0.15)}}
header h1{{font-size:18px;font-weight:600}}
header span{{font-size:12px;opacity:0.85}}
.toolbar{{background:white;padding:10px 24px;display:flex;align-items:center;
          gap:16px;border-bottom:1px solid #ddd;flex-wrap:wrap}}
.toolbar label{{font-size:12px;font-weight:600;color:#555;text-transform:uppercase;
               letter-spacing:.05em}}
.toolbar input[type=range]{{width:200px;accent-color:#1D9E75}}
.toolbar span.val{{font-size:13px;font-weight:600;color:#1D9E75;min-width:32px}}
.legend{{display:flex;gap:16px;align-items:center;margin-left:auto}}
.leg-item{{display:flex;align-items:center;gap:5px;font-size:12px}}
.leg-dot{{width:12px;height:12px;border-radius:50%;border:2px solid white;
          box-shadow:0 0 0 1px rgba(0,0,0,0.15)}}
#canvas-wrap{{overflow:auto;padding:24px;cursor:grab}}
#canvas-wrap:active{{cursor:grabbing}}
canvas{{display:block}}
.tooltip{{position:fixed;background:rgba(20,20,20,0.92);color:white;
          border-radius:10px;padding:10px 14px;font-size:12px;line-height:1.7;
          pointer-events:none;z-index:999;max-width:280px;
          box-shadow:0 4px 20px rgba(0,0,0,0.35);display:none}}
.tooltip strong{{font-size:13px;display:block;margin-bottom:3px}}
footer{{text-align:center;padding:12px;font-size:11px;color:#999;
        border-top:1px solid #e0e0e0;background:white}}
</style>
</head>
<body>
<header>
  <h1>🗺️ Timeline Visual — {label}</h1>
  <span>TTB Logistics · {hoje_str}</span>
</header>
<div class="toolbar">
  <label>Zoom</label>
  <input type="range" id="zoom" min="40" max="300" value="100"
         oninput="zoomChange(this.value)">
  <span class="val" id="zoom-val">100%</span>
  <div class="legend">
    <div class="leg-item">
      <div class="leg-dot" style="background:#1D9E75"></div> No prazo
    </div>
    <div class="leg-item">
      <div class="leg-dot" style="background:#cc2222"></div> Fora do prazo
    </div>
    <div class="leg-item">
      <div class="leg-dot" style="background:#888;border-radius:2px"></div> Armazém
    </div>
  </div>
</div>
<div id="canvas-wrap">
  <canvas id="cv"></canvas>
</div>
<div class="tooltip" id="tip"></div>
<footer>TTB Logistics AI Simulator · Timeline Visual · {label}</footer>
<script>
const DADOS   = {dados_json};
const CORES   = {cores_json};
const CORES_D = {cores_d_json};
const ORDEM   = {ordem_json};
const KM_MAX  = {km_max_global:.1f};

// Layout constantes (em px a zoom 100%)
const ROW_H   = 110;   // altura de cada linha de trip
const PAD_TOP = 60;    // margem superior
const PAD_BOT = 40;
const PAD_ESQ = 240;   // largura do label esquerdo
const PAD_DIR = 40;
const LINHA_Y = 42;    // posição Y da linha dentro da linha de trip (do topo da row)

let zoomF = 1.0;       // fator de zoom atual
let hits  = [];        // áreas clicáveis/hover para tooltip

const cv  = document.getElementById('cv');
const ctx = cv.getContext('2d');
const tip = document.getElementById('tip');

function kmToX(km) {{
  const w = cv.width;
  const usable = w - PAD_ESQ - PAD_DIR;
  return PAD_ESQ + (km / KM_MAX) * usable;
}}

function rowY(idx) {{
  return PAD_TOP + idx * ROW_H;
}}

function draw() {{
  const n  = ORDEM.length;
  const cw = Math.max(1200, window.innerWidth - 48) * zoomF;
  const ch = PAD_TOP + n * ROW_H + PAD_BOT;
  cv.width  = cw;
  cv.height = ch;
  hits = [];

  ctx.clearRect(0, 0, cw, ch);

  // Fundo
  ctx.fillStyle = '#f7f7f7';
  ctx.fillRect(0, 0, cw, ch);

  // Grade vertical (km)
  const step = gridStep();
  ctx.strokeStyle = '#ddd';
  ctx.lineWidth   = 0.8;
  ctx.setLineDash([4, 4]);
  for (let km = 0; km <= KM_MAX; km += step) {{
    const x = kmToX(km);
    ctx.beginPath();
    ctx.moveTo(x, PAD_TOP - 20);
    ctx.lineTo(x, ch - PAD_BOT + 10);
    ctx.stroke();
    ctx.fillStyle = '#999';
    ctx.font = `${{11*zoomF}}px system-ui`;
    ctx.textAlign = 'center';
    ctx.fillText(km + ' km', x, PAD_TOP - 6);
  }}
  ctx.setLineDash([]);

  // Cada trip
  ORDEM.forEach((trip, idx) => {{
    const d    = DADOS[trip];
    const cor  = CORES[trip];
    const cord = CORES_D[trip];
    const ry   = rowY(idx);
    const ly   = ry + LINHA_Y;
    const x0   = kmToX(0);
    const x1   = kmToX(d.km_tot);

    // Fundo alternado da linha
    if (idx % 2 === 0) {{
      ctx.fillStyle = 'rgba(0,0,0,0.025)';
      ctx.fillRect(0, ry, cw, ROW_H);
    }}

    // Label esquerdo
    ctx.textAlign = 'right';
    ctx.font      = `bold ${{12*zoomF}}px system-ui`;
    ctx.fillStyle = cord;
    ctx.fillText(trip + '  ' + d.cam, PAD_ESQ - 12, ly - 6);
    ctx.font      = `${{10*zoomF}}px system-ui`;
    ctx.fillStyle = '#666';
    let sub = 'Saída ' + d.saida;
    if (d.local) sub += '  ·  ' + d.local;
    sub += '  ·  ' + d.paradas.length + ' parada(s)  ·  ' + d.km_tot + ' km';
    if (d.ocup) sub += '  ·  ' + parseFloat(d.ocup).toFixed(0) + '% ocup.';
    ctx.fillText(sub, PAD_ESQ - 12, ly + 10);

    // Linha da rota
    ctx.strokeStyle = cor;
    ctx.lineWidth   = 3.5 * zoomF;
    ctx.lineCap     = 'round';
    ctx.beginPath();
    ctx.moveTo(x0, ly);
    ctx.lineTo(x1, ly);
    ctx.stroke();

    // Ponto de origem — Armazém
    drawPonto(x0, ly, '#555', 8*zoomF, false);
    ctx.font      = `italic ${{9*zoomF}}px system-ui`;
    ctx.fillStyle = '#666';
    ctx.textAlign = 'center';
    ctx.fillText('Armazém', x0, ly - 14*zoomF);
    ctx.fillText('0 km',    x0, ly - 4*zoomF);

    // Paradas
    d.paradas.forEach((p, j) => {{
      const xp  = kmToX(p.km);
      const ok  = p.prazo;
      const cpt = ok ? cord : '#cc2222';

      // Linha vertical pontilhada
      ctx.strokeStyle = cpt;
      ctx.lineWidth   = 0.8;
      ctx.setLineDash([3, 3]);
      const yBaixo = ly + (j % 2 === 0 ? 20*zoomF : 46*zoomF);
      ctx.beginPath();
      ctx.moveTo(xp, ly + 6*zoomF);
      ctx.lineTo(xp, yBaixo + 2);
      ctx.stroke();
      ctx.setLineDash([]);

      // KM acima
      ctx.font      = `${{9*zoomF}}px system-ui`;
      ctx.fillStyle = '#555';
      ctx.textAlign = 'center';
      ctx.fillText(p.km + ' km', xp, ly - 16*zoomF);

      // Ponto
      drawPonto(xp, ly, cpt, 7*zoomF, !ok);

      // Box com nome e horário abaixo
      const linhas = [
        p.cli.length > 24 ? p.cli.substring(0,23)+'…' : p.cli,
        p.chegada + '  (lim ' + p.limite + ')'
      ];
      const bx = xp;
      const by = yBaixo + 4;
      drawBox(bx, by, linhas, cpt, 9*zoomF);

      // Área para tooltip
      hits.push({{
        x: xp - 14*zoomF, y: ly - 14*zoomF,
        w: 28*zoomF,       h: 28*zoomF,
        trip, j, p
      }});
    }});
  }});
}}

function drawPonto(x, y, cor, r, cross) {{
  ctx.beginPath();
  ctx.arc(x, y, r, 0, Math.PI * 2);
  ctx.fillStyle = cor;
  ctx.fill();
  ctx.strokeStyle = 'white';
  ctx.lineWidth = 1.8;
  ctx.stroke();
  if (cross) {{
    ctx.strokeStyle = 'white';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(x-r*.5, y-r*.5); ctx.lineTo(x+r*.5, y+r*.5);
    ctx.moveTo(x+r*.5, y-r*.5); ctx.lineTo(x-r*.5, y+r*.5);
    ctx.stroke();
  }}
}}

function drawBox(cx, by, linhas, cor, fs) {{
  ctx.font = fs + 'px system-ui';
  const maxW = Math.max(...linhas.map(l => ctx.measureText(l).width));
  const bw = maxW + 14;
  const bh = linhas.length * (fs + 4) + 8;
  const bx = cx - bw/2;

  ctx.fillStyle = 'white';
  roundRect(bx, by, bw, bh, 5);
  ctx.fill();
  ctx.strokeStyle = cor;
  ctx.lineWidth = 1;
  roundRect(bx, by, bw, bh, 5);
  ctx.stroke();

  ctx.fillStyle = cor;
  ctx.textAlign = 'center';
  linhas.forEach((l, i) => {{
    ctx.fillText(l, cx, by + (fs + 4)*(i+1));
  }});
}}

function roundRect(x, y, w, h, r) {{
  ctx.beginPath();
  ctx.moveTo(x+r, y);
  ctx.lineTo(x+w-r, y);
  ctx.quadraticCurveTo(x+w, y, x+w, y+r);
  ctx.lineTo(x+w, y+h-r);
  ctx.quadraticCurveTo(x+w, y+h, x+w-r, y+h);
  ctx.lineTo(x+r, y+h);
  ctx.quadraticCurveTo(x, y+h, x, y+h-r);
  ctx.lineTo(x, y+r);
  ctx.quadraticCurveTo(x, y, x+r, y);
  ctx.closePath();
}}

function gridStep() {{
  const steps = [5,10,20,25,50,100,150,200,250,500];
  const ideal = KM_MAX / 10;
  return steps.find(s => s >= ideal) || 500;
}}

function zoomChange(v) {{
  zoomF = v / 100;
  document.getElementById('zoom-val').textContent = v + '%';
  draw();
}}

// Tooltip ao passar o mouse
cv.addEventListener('mousemove', e => {{
  const rect = cv.getBoundingClientRect();
  const mx = (e.clientX - rect.left) * (cv.width / rect.width);
  const my = (e.clientY - rect.top)  * (cv.height / rect.height);
  let found = null;
  for (const h of hits) {{
    if (mx >= h.x && mx <= h.x+h.w && my >= h.y && my <= h.y+h.h) {{
      found = h; break;
    }}
  }}
  if (found) {{
    const p = found.p;
    const d = DADOS[found.trip];
    tip.style.display = 'block';
    tip.style.left    = (e.clientX + 16) + 'px';
    tip.style.top     = (e.clientY - 10) + 'px';
    const prazoTxt = p.prazo ? '✅ No prazo' : '❌ Fora do prazo';
    const descTxt  = p.descarga ? `<br>⏱ Descarga: ${{p.descarga}} min` : '';
    const nfTxt    = p.nfs ? `<br>📄 NF(s): ${{p.nfs}}` : '';
    tip.innerHTML = `<strong>${{p.cli}}</strong>
      🕐 Chegada: ${{p.chegada}}<br>
      ⏰ Limite: ${{p.limite}}<br>
      ${{prazoTxt}}${{descTxt}}${{nfTxt}}<br>
      📍 KM acumulado: ${{p.km}} km`;
    cv.style.cursor = 'pointer';
  }} else {{
    tip.style.display = 'none';
    cv.style.cursor = 'grab';
  }}
}});

cv.addEventListener('mouseleave', () => {{ tip.style.display='none'; }});

window.addEventListener('resize', draw);
draw();
</script>
</body>
</html>"""

    with open(caminho_saida, 'w', encoding='utf-8') as f:
        f.write(html)

    salvar_e_abrir(caminho_saida, silent=True)
    print(f"   🗺️  Timeline HTML salvo e aberto → {os.path.basename(caminho_saida)}")


print("\n🗺️  Gerando Timelines Visuais HTML...")
gerar_timeline_visual_html(df_tl,  'Plano Real',
                           os.path.join(DRIVE_PATH, 'Timeline_Visual_REAL.html'))
gerar_timeline_visual_html(df_otm, 'Plano Otimizado',
                           os.path.join(DRIVE_PATH, 'Timeline_Visual_OTM.html'))
if 'df_ag' in globals() and isinstance(df_ag, pd.DataFrame) and not df_ag.empty:
    gerar_timeline_visual_html(df_ag, 'Plano AG (Algoritmo Genético)',
                               os.path.join(DRIVE_PATH, 'Timeline_Visual_AG.html'))
else:
    print("   ⚠️  AG: df_ag não disponível — Timeline Visual AG ignorado.")

# ── Seção interestadual nos timelines ────────────────────────
# Adiciona bloco HTML de entregas interestaduais ao final de cada timeline.
if ('df_interestadual' in globals() and
        isinstance(df_interestadual, pd.DataFrame) and
        not df_interestadual.empty):

    _INTER_CORES = {
        'Sul':          '#f97316',   # laranja
        'Sudeste':      '#a855f7',   # roxo
        'Centro-Oeste': '#06b6d4',   # ciano
        'Norte':        '#84cc16',   # verde-lima
        'Nordeste':     '#f43f5e',   # rosa
    }

    def _bloco_inter_html(df_inter):
        """Gera bloco HTML da tabela interestadual para injetar nos timelines."""
        linhas = ''
        for _, row in df_inter.iterrows():
            _reg  = row.get('Região', '')
            _cor  = _INTER_CORES.get(_reg, '#94a3b8')
            _uf   = row.get('UF', '')
            _cli  = row.get('Cliente', '')
            _cam  = row.get('Caminhão', '')
            _eng  = row.get('Qtd Engradados', 0)
            _ocp  = row.get('Ocupação %', 0)
            _dist = row.get('Distancia KM (est)', 0)
            _fr   = row.get('Frete Estimado (R$)', 0)
            _ufs  = row.get('UFs na Viagem', '')
            _via  = row.get('Viagem', '')
            linhas += (
                f'<tr>'
                f'<td><span style="background:{_cor};color:white;border-radius:4px;'
                f'padding:2px 7px;font-size:11px;font-weight:600">{_via}</span></td>'
                f'<td style="color:{_cor};font-weight:600">{_reg}</td>'
                f'<td style="font-weight:700">{_uf}</td>'
                f'<td>{_cli}</td>'
                f'<td style="color:#94a3b8">{_cam}</td>'
                f'<td style="text-align:right">{int(_eng)}</td>'
                f'<td style="text-align:right">{_ocp:.1f}%</td>'
                f'<td style="text-align:right">{_dist:.0f} km</td>'
                f'<td style="text-align:right;color:#fbbf24;font-weight:600">'
                f'R$ {_fr:,.2f}</td>'
                f'<td style="color:#64748b;font-size:11px">{_ufs}</td>'
                f'</tr>'
            )
        _n_trips = df_inter['Viagem'].nunique()
        _fr_tot  = df_inter['Frete Estimado (R$)'].sum()
        _regioes = ', '.join(sorted(df_inter['Região'].unique()))
        return f"""
<div style="margin:32px 16px 16px;background:#1e293b;border:1px solid #f97316;
            border-radius:14px;overflow:hidden;font-family:sans-serif">
  <div style="background:#1a1200;padding:14px 20px;border-bottom:1px solid #f97316;
              display:flex;align-items:center;gap:16px;flex-wrap:wrap">
    <span style="font-size:18px">🚛</span>
    <div>
      <div style="font-size:15px;font-weight:700;color:#fbbf24">
        Entregas Interestaduais — Transportadora Terceirizada (rodovia)</div>
      <div style="font-size:12px;color:#92400e;margin-top:2px">
        {_n_trips} trip(s) &middot; Regiões: {_regioes}
        &middot; Frete est.: R$ {_fr_tot:,.2f}
        &middot; Linha tracejada no mapa = rota estimada</div>
    </div>
  </div>
  <div style="overflow-x:auto">
  <table style="width:100%;border-collapse:collapse;font-size:13px">
    <thead>
      <tr style="background:#162032">
        <th style="padding:10px 14px;text-align:left;color:#fbbf24;font-size:11px;
                   text-transform:uppercase;letter-spacing:.05em">Viagem</th>
        <th style="padding:10px 14px;text-align:left;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Região</th>
        <th style="padding:10px 14px;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">UF</th>
        <th style="padding:10px 14px;text-align:left;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Cliente</th>
        <th style="padding:10px 14px;text-align:left;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Caminhão</th>
        <th style="padding:10px 14px;text-align:right;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Eng.</th>
        <th style="padding:10px 14px;text-align:right;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Ocup.</th>
        <th style="padding:10px 14px;text-align:right;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Dist (km)</th>
        <th style="padding:10px 14px;text-align:right;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">Frete Est.</th>
        <th style="padding:10px 14px;text-align:left;color:#fbbf24;font-size:11px;
                   text-transform:uppercase">UFs Trip</th>
      </tr>
    </thead>
    <tbody style="color:#cbd5e1">
      {linhas}
    </tbody>
  </table>
  </div>
</div>"""

    _bloco_html = _bloco_inter_html(df_interestadual)

    # Injeta o bloco em cada timeline já gerado
    for _tl_path in [
        os.path.join(DRIVE_PATH, 'Timeline_Visual_REAL.html'),
        os.path.join(DRIVE_PATH, 'Timeline_Visual_OTM.html'),
        os.path.join(DRIVE_PATH, 'Timeline_Visual_AG.html'),
    ]:
        if not os.path.exists(_tl_path):
            continue
        try:
            with open(_tl_path, 'r', encoding='utf-8') as _f:
                _conteudo = _f.read()
            _conteudo = _conteudo.replace('</body>', _bloco_html + '\n</body>')
            with open(_tl_path, 'w', encoding='utf-8') as _f:
                _f.write(_conteudo)
            print(f"   ✈️  Interestaduais injetados → {os.path.basename(_tl_path)}")
        except Exception as _e_inj:
            print(f"   ⚠️  Erro ao injetar interestadual em {os.path.basename(_tl_path)}: {repr(_e_inj)}")

# =============================================================
# MAPA AG — Folium com polylines idênticas ao mapa otimizado
# =============================================================
if 'df_ag' in globals() and not df_ag.empty:
    print("\n🗺️  Gerando mapa AG...")
    try:
        import folium

        n_viag_ag_mapa = df_ag['Viagem'].nunique()
        mapa_ag = folium.Map(
            location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
            zoom_start=9, tiles='OpenStreetMap'
        )

        # Título
        titulo_ag = (
            f'<div style="position:fixed;top:10px;left:50%;transform:translateX(-50%);'
            f'z-index:1000;background:white;padding:10px 20px;border-radius:8px;'
            f'border:2px solid #7F77DD;font-size:13px;font-weight:bold;'
            f'box-shadow:2px 2px 6px rgba(0,0,0,.3);text-align:center">'
            f'Plano AG — Algoritmo Genético<br>'
            f'<span style="font-weight:normal;font-size:11px">'
            f'{n_viag_ag_mapa} viagens</span></div>'
        )
        mapa_ag.get_root().html.add_child(folium.Element(titulo_ag))

        # Armazém
        folium.Marker(
            location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
            popup=folium.Popup('<b>Armazém Suzano</b>', max_width=200),
            tooltip='Armazém Suzano',
            icon=folium.Icon(color='black', icon='home', prefix='fa')
        ).add_to(mapa_ag)

        # Agrupa por viagem
        viagens_ag_map = {}
        for r in timeline_ag:
            viagens_ag_map.setdefault(r['Viagem'], []).append(r)

        for idx_v, (trip_id, paradas) in enumerate(sorted(viagens_ag_map.items())):
            cor = CORES_VIAGEM[idx_v % len(CORES_VIAGEM)]

            # Monta lista de pontos: armazém → clientes em ordem
            pontos = [(ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])]
            for r in paradas:
                coord = coord_cli.get(r['Cliente'])
                if coord:
                    pontos.append(coord)

            # Busca polyline real (cache → API → linha reta como fallback)
            rota_real = []
            if len(pontos) > 1:
                chave_dir = str(sorted([str(p) for p in pontos]))
                if chave_dir in _cache_directions:
                    rota_real = _cache_directions[chave_dir]
                elif API_DISPONIVEL and not controlador_api.bloqueado:
                    try:
                        n_elem_dir = max(1, len(pontos) - 1)
                        if controlador_api.pode_chamar(n_elem_dir):
                            dir_res = gmaps.directions(
                                origin=pontos[0],
                                destination=pontos[-1],
                                waypoints=pontos[1:-1] if len(pontos) > 2 else None,
                                mode='driving'
                            )
                            if dir_res:
                                import polyline as polyline_lib
                                for leg in dir_res[0]['legs']:
                                    for step in leg['steps']:
                                        pts = polyline_lib.decode(
                                            step['polyline']['points'])
                                        rota_real.extend(pts)
                                controlador_api.registrar(1, n_elem_dir,
                                                          f'dir_ag_{trip_id}')
                                _cache_directions[chave_dir] = rota_real
                    except Exception:
                        pass

                folium.PolyLine(
                    locations=rota_real if rota_real else pontos,
                    color=cor, weight=4, opacity=0.85,
                    tooltip=f"{trip_id} — {len(paradas)} cliente(s)"
                            + ('' if rota_real else ' (linha reta)')
                ).add_to(mapa_ag)

            # Marcadores dos clientes
            for r in paradas:
                coord = coord_cli.get(r['Cliente'])
                if not coord:
                    continue
                cli_nome  = str(r['Cliente']).split('[')[0].strip()
                prazo_txt = r.get('Dentro do Prazo', '?')
                ok        = '✅' in prazo_txt
                popup_html = (
                    f"<b>{cli_nome}</b><br>"
                    f"Viagem: {trip_id} · {r.get('Caminhão','')}<br>"
                    f"NFs: {r.get('NFs','—')}<br>"
                    f"Chegada: {r.get('Chegada','—')} "
                    f"(limite {str(r.get('Limite Entrega','—')).replace(':00','').strip()})<br>"
                    f"{prazo_txt}"
                )
                folium.CircleMarker(
                    location=coord, radius=8,
                    color=cor, fill=True,
                    fill_color='white' if ok else '#ff4444',
                    fill_opacity=0.9, weight=2.5,
                    popup=folium.Popup(popup_html, max_width=260),
                    tooltip=f"{trip_id} · {cli_nome}"
                ).add_to(mapa_ag)

        # Legenda
        legenda_ag = (
            '<div style="position:fixed;bottom:30px;right:10px;z-index:1000;'
            'background:white;padding:10px 14px;border-radius:8px;'
            'border:1px solid #ccc;font-size:12px;box-shadow:2px 2px 6px rgba(0,0,0,.2)">'
            '<b>Plano AG</b><br>'
            '<span style="color:#1D9E75">&#9679;</span> No prazo<br>'
            '<span style="color:#ff4444">&#9679;</span> Fora do prazo'
            '</div>'
        )
        mapa_ag.get_root().html.add_child(folium.Element(legenda_ag))

        caminho_mapa_ag = os.path.join(DRIVE_PATH, 'mapa_rotas_ag.html')
        mapa_ag.save(caminho_mapa_ag)
        salvar_e_abrir(caminho_mapa_ag, silent=True)
        print(f"   ✅ Mapa AG salvo → {caminho_mapa_ag}")

    except ImportError:
        print("   ⚠️  folium não instalado.")
    except Exception as e_ag_mapa:
        print(f"   ⚠️  Erro ao gerar mapa AG: {repr(e_ag_mapa)}")

# =============================================================
# MAPA DBSCAN+AG — mesmo estilo do mapa Original
# (trips coloridas + contagem de paradas na legenda)
# =============================================================
if 'df_dbscan_ag' in globals() and not df_dbscan_ag.empty:
    print("\n🗺️  Gerando mapa DBSCAN+AG...")
    try:
        import folium

        n_viag_dba_mapa = df_dbscan_ag['Viagem'].nunique()
        mapa_dba = folium.Map(
            location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
            zoom_start=9, tiles='OpenStreetMap'
        )

        # Título fixo no topo
        titulo_dba = (
            f'<div style="position:fixed;top:10px;left:50%;transform:translateX(-50%);'
            f'z-index:1000;background:white;padding:10px 20px;border-radius:8px;'
            f'border:2px solid #22d3ee;font-size:13px;font-weight:bold;'
            f'box-shadow:2px 2px 6px rgba(0,0,0,.3);text-align:center">'
            f'Plano DBSCAN+AG — Híbrido Geográfico + Genético<br>'
            f'<span style="font-weight:normal;font-size:11px">'
            f'{n_viag_dba_mapa} viagem(ns)</span></div>'
        )
        mapa_dba.get_root().html.add_child(folium.Element(titulo_dba))

        # Marcador do armazém
        folium.Marker(
            location=[ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon']],
            popup=folium.Popup('<b>Armazém Suzano</b>', max_width=200),
            tooltip='Armazém Suzano',
            icon=folium.Icon(color='black', icon='home', prefix='fa')
        ).add_to(mapa_dba)

        # Agrupa paradas por viagem usando timeline_dbscan_ag
        viagens_dba_map = {}
        for r in timeline_dbscan_ag:
            viagens_dba_map.setdefault(r['Viagem'], []).append(r)

        for idx_v, (trip_id, paradas) in enumerate(sorted(viagens_dba_map.items())):
            cor = CORES_VIAGEM[idx_v % len(CORES_VIAGEM)]

            # Linha da rota: armazém → clientes em ordem de entrega
            pontos = [(ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])]
            for r in paradas:
                coord = coord_cli.get(r['Cliente'])
                if coord:
                    pontos.append(coord)

            # Polyline: cache → API Google Directions → linha reta (fallback)
            rota_real = []
            if len(pontos) > 1:
                chave_dir = str(sorted([str(p) for p in pontos]))
                if chave_dir in _cache_directions:
                    rota_real = _cache_directions[chave_dir]
                elif API_DISPONIVEL and not controlador_api.bloqueado:
                    try:
                        n_elem_dir = max(1, len(pontos) - 1)
                        if controlador_api.pode_chamar(n_elem_dir):
                            dir_res = gmaps.directions(
                                origin=pontos[0],
                                destination=pontos[-1],
                                waypoints=pontos[1:-1] if len(pontos) > 2 else None,
                                mode='driving'
                            )
                            if dir_res:
                                import polyline as polyline_lib
                                for leg in dir_res[0]['legs']:
                                    for step in leg['steps']:
                                        pts = polyline_lib.decode(
                                            step['polyline']['points'])
                                        rota_real.extend(pts)
                                controlador_api.registrar(1, n_elem_dir,
                                                          f'dir_dba_{trip_id}')
                                _cache_directions[chave_dir] = rota_real
                    except Exception:
                        pass

                folium.PolyLine(
                    locations=rota_real if rota_real else pontos,
                    color=cor, weight=4, opacity=0.85,
                    tooltip=f"{trip_id} — {len(paradas)} parada(s)"
                            + ('' if rota_real else ' (linha reta)')
                ).add_to(mapa_dba)

            # Marcadores dos clientes
            for r in paradas:
                coord = coord_cli.get(r['Cliente'])
                if not coord:
                    continue
                cli_nome  = str(r['Cliente']).split('[')[0].strip()
                prazo_txt = r.get('Dentro do Prazo', '?')
                ok        = '✅' in prazo_txt
                popup_html = (
                    f"<b>{cli_nome}</b><br>"
                    f"Viagem: {trip_id} · {r.get('Caminhão','')}<br>"
                    f"NFs: {r.get('NFs','—')}<br>"
                    f"Chegada: {r.get('Chegada','—')} "
                    f"(limite {str(r.get('Limite Entrega','—')).replace(':00','').strip()})<br>"
                    f"{prazo_txt}"
                )
                folium.CircleMarker(
                    location=coord, radius=8,
                    color=cor, fill=True,
                    fill_color='white' if ok else '#ff4444',
                    fill_opacity=0.9, weight=2.5,
                    popup=folium.Popup(popup_html, max_width=260),
                    tooltip=f"{trip_id} · {cli_nome}"
                ).add_to(mapa_dba)

        # Legenda: trips com contagem de paradas (igual ao mapa Original)
        linhas_leg_dba = ''.join(
            f'<span style="color:{CORES_VIAGEM[i % len(CORES_VIAGEM)]}">&#9644;</span> '
            f'{trip} ({len(viagens_dba_map[trip])} parada'
            f'{"s" if len(viagens_dba_map[trip]) > 1 else ""})<br>'
            for i, trip in enumerate(sorted(viagens_dba_map.keys()))
        )
        legenda_dba = (
            '<div style="position:fixed;bottom:30px;left:30px;z-index:1000;'
            'background:white;padding:12px;border-radius:8px;border:1px solid #22d3ee;'
            'font-size:12px;max-height:320px;overflow-y:auto;'
            'box-shadow:2px 2px 6px rgba(0,0,0,.3)">'
            f'<b>Plano DBSCAN+AG</b><br>{linhas_leg_dba}'
            '<br><span style="color:#1D9E75">&#9679;</span> No prazo&nbsp;&nbsp;'
            '<span style="color:#ff4444">&#9679;</span> Fora do prazo'
            '</div>'
        )
        mapa_dba.get_root().html.add_child(folium.Element(legenda_dba))

        caminho_mapa_dba = os.path.join(DRIVE_PATH, 'mapa_rotas_dbscan_ag.html')
        mapa_dba.save(caminho_mapa_dba)
        salvar_e_abrir(caminho_mapa_dba, silent=True)
        print(f"   ✅ Mapa DBSCAN+AG salvo → {caminho_mapa_dba}")

    except ImportError:
        print("   ⚠️  folium não instalado.")
    except Exception as e_dba_mapa:
        print(f"   ⚠️  Erro ao gerar mapa DBSCAN+AG: {repr(e_dba_mapa)}")

# Persiste cache de directions para evitar chamadas repetidas à API
try:
    import pickle
    with open(DIRECTIONS_CACHE_FILE, 'wb') as _f:
        pickle.dump(_cache_directions, _f)
    print(f"   💾 Cache de polylines salvo: {len(_cache_directions)} rotas")
except Exception as e_pkl:
    print(f"   ⚠️  Erro ao salvar cache directions: {repr(e_pkl)}")

controlador_api.resumo()

# =============================================================
# SUMÁRIO HTML — comparativo completo gerado ao final
# =============================================================
def gerar_html_sumario(caminho_saida):
    """
    Gera HTML com o sumário comparativo de até 5 planos:
    Original | AG | DBSCAN | DBSCAN+AG | RL (DQN)
    """
    import json
    from datetime import datetime as _dt

    hoje = _dt.now().strftime('%d/%m/%Y %H:%M')

    # ── Coleta dados ──────────────────────────────────────────
    _df_ag_ok  = 'df_ag'       in globals() and isinstance(df_ag,       pd.DataFrame) and not df_ag.empty
    _df_db_ok  = 'df_dbscan'   in globals() and isinstance(df_dbscan,   pd.DataFrame) and not df_dbscan.empty
    _df_dba_ok = 'df_dbscan_ag'in globals() and isinstance(df_dbscan_ag,pd.DataFrame) and not df_dbscan_ag.empty
    _df_rl_ok  = 'df_rl'       in globals() and isinstance(df_rl,       pd.DataFrame) and not df_rl.empty
    tem_ag  = _df_ag_ok
    tem_db  = _df_db_ok
    tem_dba = _df_dba_ok
    tem_rl  = _df_rl_ok

    df_ag_s   = df_ag        if tem_ag  else pd.DataFrame()
    df_db_s   = df_dbscan    if tem_db  else pd.DataFrame()
    df_dba_s  = df_dbscan_ag if tem_dba else pd.DataFrame()
    df_rl_s   = df_rl        if tem_rl  else pd.DataFrame()

    def _col_cam(df):
        return 'Caminhão' if 'Caminhão' in df.columns else 'Caminhao'
    def _n_viag(df):
        return df['Viagem'].nunique() if not df.empty else 0
    def _ocup(df):
        return round(df['Ocupação %'].mean(), 1) if (not df.empty and 'Ocupação %' in df.columns) else 0
    def _km(df):
        return round(df['Distancia KM'].sum(), 1) if (not df.empty and 'Distancia KM' in df.columns) else 0
    def _prazo(df):
        if df.empty or 'Dentro do Prazo' not in df.columns: return 0, 0
        return int((df['Dentro do Prazo'] == '✅').sum()), len(df)
    def _cli_viag(df):
        if df.empty or 'Viagem' not in df.columns: return 0
        return round(df.groupby('Viagem')['Cliente'].count().mean(), 1)
    def _tipos(df):
        col = _col_cam(df)
        if df.empty or col not in df.columns: return {}
        return df.groupby('Viagem')[col].first().value_counts().to_dict()
    def _co2_total(df):
        return round(df['CO2 (kg)'].sum(), 0) if (not df.empty and 'CO2 (kg)' in df.columns) else 0
    def _co2_tipo(df):
        col = _col_cam(df)
        if df.empty or 'CO2 (kg)' not in df.columns: return {}
        return {str(k): round(v, 0) for k, v in df.groupby(col)['CO2 (kg)'].sum().items()}
        # frete total = soma por viagem (não por parada)
        return round(df.groupby('Viagem')['Frete (R$)'].first().sum() *
                     df.groupby('Viagem').size().mean() /
                     df.groupby('Viagem').size().mean(), 2)

    # Usa variáveis já calculadas no script quando disponíveis
    fr_real  = globals().get('frete_real',      0)
    fr_otm   = globals().get('frete_otm',       0)
    fr_ag    = globals().get('frete_ag',        0) if tem_ag  else 0
    fr_db    = globals().get('frete_dbscan',    0) if tem_db  else 0
    fr_dba   = globals().get('frete_dbscan_ag', 0) if tem_dba else 0
    fr_rl    = globals().get('frete_rl',        0) if tem_rl  else 0

    n_real = _n_viag(df_tl);    n_otm  = _n_viag(df_otm)
    n_ag   = _n_viag(df_ag_s);  n_db   = _n_viag(df_db_s)
    n_dba  = _n_viag(df_dba_s); n_rl   = _n_viag(df_rl_s)

    km_real = _km(df_tl);    km_otm  = _km(df_otm)
    km_ag   = _km(df_ag_s);  km_db   = _km(df_db_s)
    km_dba  = _km(df_dba_s); km_rl   = _km(df_rl_s)

    oc_real = _ocup(df_tl);    oc_otm  = _ocup(df_otm)
    oc_ag   = _ocup(df_ag_s);  oc_db   = _ocup(df_db_s)
    oc_dba  = _ocup(df_dba_s); oc_rl   = _ocup(df_rl_s)

    cv_real = _cli_viag(df_tl);    cv_otm  = _cli_viag(df_otm)
    cv_ag   = _cli_viag(df_ag_s);  cv_db   = _cli_viag(df_db_s)
    cv_dba  = _cli_viag(df_dba_s); cv_rl   = _cli_viag(df_rl_s)

    pr_ok_real, pr_tot_real = _prazo(df_tl)
    pr_ok_otm,  pr_tot_otm  = _prazo(df_otm)
    pr_ok_ag,   pr_tot_ag   = (_prazo(df_ag_s)  if tem_ag  else (0, 0))
    pr_ok_db,   pr_tot_db   = (_prazo(df_db_s)  if tem_db  else (0, 0))
    pr_ok_dba,  pr_tot_dba  = (_prazo(df_dba_s) if tem_dba else (0, 0))
    pr_ok_rl,   pr_tot_rl   = (_prazo(df_rl_s)  if tem_rl  else (0, 0))

    co2_real = globals().get('co2_real_kg',      _co2_total(df_tl))
    co2_otm  = globals().get('co2_otm_kg',       _co2_total(df_otm))
    co2_ag   = globals().get('co2_ag_kg',        _co2_total(df_ag_s)  if tem_ag  else 0)
    co2_db   = globals().get('co2_dbscan_kg',    _co2_total(df_db_s)  if tem_db  else 0)
    co2_dba  = globals().get('co2_dbscan_ag_kg', _co2_total(df_dba_s) if tem_dba else 0)
    co2_rl   = globals().get('co2_rl_kg',        _co2_total(df_rl_s)  if tem_rl  else 0)

    tipos_real = _tipos(df_tl);    tipos_otm  = _tipos(df_otm)
    tipos_ag   = _tipos(df_ag_s);  tipos_db   = _tipos(df_db_s)
    tipos_dba  = _tipos(df_dba_s); tipos_rl   = _tipos(df_rl_s)
    all_tipos  = sorted(set(list(tipos_real)+list(tipos_ag)+list(tipos_db)+list(tipos_dba)+list(tipos_rl)))

    co2t_real = _co2_tipo(df_tl);    co2t_otm  = _co2_tipo(df_otm)
    co2t_ag   = _co2_tipo(df_ag_s);  co2t_db   = _co2_tipo(df_db_s)
    co2t_dba  = _co2_tipo(df_dba_s); co2t_rl   = _co2_tipo(df_rl_s)
    all_co2t  = sorted(set(list(co2t_real)+list(co2t_ag)+list(co2t_db)+list(co2t_dba)+list(co2t_rl)))

    def pct(base, novo):
        if base <= 0 or novo == 0: return ''
        return f"+{(base-novo)/base*100:.1f}%"

    def fmt_r(v):
        return f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    kpis = [
        {'label': 'Viagens',          'r': n_real,   'o': n_otm,  'a': n_ag,   'db': n_db,   'dba': n_dba,  'rl': n_rl,   'fmt': 'd',    'inv': True},
        {'label': 'Ocupação média',   'r': oc_real,  'o': oc_otm, 'a': oc_ag,  'db': oc_db,  'dba': oc_dba, 'rl': oc_rl,  'fmt': '.1f%', 'inv': False},
        {'label': 'Distância (km)',   'r': km_real,  'o': km_otm, 'a': km_ag,  'db': km_db,  'dba': km_dba, 'rl': km_rl,  'fmt': '.0f',  'inv': True},
        {'label': 'Frete total',      'r': fr_real,  'o': fr_otm, 'a': fr_ag,  'db': fr_db,  'dba': fr_dba, 'rl': fr_rl,  'fmt': 'brl',  'inv': True},
        {'label': 'Emissão CO₂ (kg)','r': co2_real, 'o': co2_otm,'a': co2_ag, 'db': co2_db, 'dba': co2_dba,'rl': co2_rl, 'fmt': '.0f',  'inv': True},
        {'label': 'Clientes/viagem',  'r': cv_real,  'o': cv_otm, 'a': cv_ag,  'db': cv_db,  'dba': cv_dba, 'rl': cv_rl,  'fmt': '.1f',  'inv': False},
    ]

    def fmt_val(v, f):
        if f == 'd':    return str(int(v))
        if f == '.1f%': return f"{v:.1f}%"
        if f == '.0f':  return f"{v:.0f}"
        if f == '.1f':  return f"{v:.1f}"
        if f == 'brl':  return fmt_r(v)
        return str(v)

    def delta_class(base, novo, inv):
        if base <= 0 or novo == 0: return 'neutral'
        if inv:  return 'good' if novo < base else 'bad'
        else:    return 'good' if novo > base else 'bad'

    planos_ativos = [('Original', None, None, 'red')]
    if tem_ag:  planos_ativos += [('AG',        'a',   None, 'purple')]
    if tem_db:  planos_ativos += [('DBSCAN',    'db',  None, 'amber')]
    if tem_dba: planos_ativos += [('DBSCAN+AG', 'dba', None, 'cyan')]
    if tem_rl:  planos_ativos += [('RL (DQN)',  'rl',  None, 'teal')]

    kpi_html = ''
    for k in kpis:
        r, f, inv = k['r'], k['fmt'], k['inv']
        kpi_html += f'<div class="kpi-card"><div class="kpi-label">{k["label"]}</div><div class="kpi-row">'
        for nome, key, _, cor in planos_ativos:
            val   = k[key] if key else r
            pct_v = pct(r, val) if (key and inv) else (pct(val, r) if key else '')
            cls_v = delta_class(r, val, inv) if key else 'neutral'
            kpi_html += (f'<div class="kpi-col">'
                         f'<div class="kpi-head">{nome}</div>'
                         f'<div class="kpi-val {cor if not key else cor}">{fmt_val(val, f)}</div>'
                         f'{"<div class=\"kpi-delta " + cls_v + "\">" + pct_v + "</div>" if key and pct_v else ""}'
                         f'</div>')
        kpi_html += '</div></div>'

    # Build detail table
    n_extra = (1 if tem_ag else 0) + (1 if tem_db else 0) + (1 if tem_dba else 0)
    ag_col_header  = '<th>AG</th>'        if tem_ag  else ''
    db_col_header  = '<th>DBSCAN</th>'    if tem_db  else ''
    dba_col_header = '<th>DBSCAN+AG</th>' if tem_dba else ''

    def tr(label, vr, va='', vdb='', vdba='', vrl='', indent=False, bold=False, sep=False):
        cls = 'tr-sep' if sep else ('tr-sub' if indent else '')
        bld = 'font-weight:600;' if bold else ''
        pfx = '&nbsp;&nbsp;&nbsp;└&nbsp;' if indent else ''
        row = (f'<tr class="{cls}"><td style="{bld}">{pfx}{label}</td>'
               f'<td class="num">{vr}</td>')
        if tem_ag:  row += f'<td class="num">{va}</td>'
        if tem_db:  row += f'<td class="num">{vdb}</td>'
        if tem_dba: row += f'<td class="num">{vdba}</td>'
        if tem_rl:  row += f'<td class="num">{vrl}</td>'
        return row + '</tr>'

    def tr_pct(label, base, *novos_pares):
        row = f'<tr><td style="color:#64748b">{label}</td><td class="num" style="color:#64748b">—</td>'
        for val, tem in novos_pares:
            if tem:
                p = pct(base, val)
                row += f'<td class="num pct-good">{p}</td>'
        return row + '</tr>'

    rows = ''
    rows += tr('Viagens (caminhões)', n_real,
               n_ag if tem_ag else '', n_db if tem_db else '',
               n_dba if tem_dba else '', n_rl if tem_rl else '', bold=True)
    for tp in all_tipos:
        rows += tr(tp, tipos_real.get(tp,0),
                   tipos_ag.get(tp,0)  if tem_ag  else '',
                   tipos_db.get(tp,0)  if tem_db  else '',
                   tipos_dba.get(tp,0) if tem_dba else '',
                   tipos_rl.get(tp,0)  if tem_rl  else '', indent=True)
    rows += tr('','','',sep=True)
    rows += tr('Clientes/viagem',     f'{cv_real:.1f}',
               f'{cv_ag:.1f}'  if tem_ag  else '', f'{cv_db:.1f}'  if tem_db  else '',
               f'{cv_dba:.1f}' if tem_dba else '', f'{cv_rl:.1f}'  if tem_rl  else '')
    rows += tr('Ocupação média (%)',  f'{oc_real:.1f}',
               f'{oc_ag:.1f}'  if tem_ag  else '', f'{oc_db:.1f}'  if tem_db  else '',
               f'{oc_dba:.1f}' if tem_dba else '', f'{oc_rl:.1f}'  if tem_rl  else '')
    rows += tr('Distância total (km)',f'{km_real:.1f}',
               f'{km_ag:.1f}'  if tem_ag  else '', f'{km_db:.1f}'  if tem_db  else '',
               f'{km_dba:.1f}' if tem_dba else '', f'{km_rl:.1f}'  if tem_rl  else '')
    rows += tr_pct('Redução distância (%)', km_real,
                   (km_ag, tem_ag), (km_db, tem_db),
                   (km_dba, tem_dba), (km_rl, tem_rl))
    if fr_real > 0:
        rows += tr('Frete total (R$)', fmt_r(fr_real),
                   fmt_r(fr_ag)  if tem_ag  else '', fmt_r(fr_db)  if tem_db  else '',
                   fmt_r(fr_dba) if tem_dba else '', fmt_r(fr_rl)  if tem_rl  else '')
        rows += tr_pct('Redução frete (%)', fr_real,
                       (fr_ag, tem_ag), (fr_db, tem_db),
                       (fr_dba, tem_dba), (fr_rl, tem_rl))
    rows += tr('','','',sep=True)
    rows += tr('Emissão CO₂ (kg)',   f'{co2_real:.0f}',
               f'{co2_ag:.0f}'  if tem_ag  else '', f'{co2_db:.0f}'  if tem_db  else '',
               f'{co2_dba:.0f}' if tem_dba else '', f'{co2_rl:.0f}'  if tem_rl  else '', bold=True)
    rows += tr_pct('Redução CO₂ (%)', co2_real,
                   (co2_ag, tem_ag), (co2_db, tem_db),
                   (co2_dba, tem_dba), (co2_rl, tem_rl))
    for tp in all_co2t:
        rows += tr(tp, f'{co2t_real.get(tp,0):.0f}',
                   f'{co2t_ag.get(tp,0):.0f}'  if tem_ag  else '',
                   f'{co2t_db.get(tp,0):.0f}'  if tem_db  else '',
                   f'{co2t_dba.get(tp,0):.0f}' if tem_dba else '',
                   f'{co2t_rl.get(tp,0):.0f}'  if tem_rl  else '', indent=True)
    rows += tr('','','',sep=True)
    rows += tr('Entregas no prazo',
               f'{pr_ok_real}/{pr_tot_real}',
               f'{pr_ok_ag}/{pr_tot_ag}'   if tem_ag  else '',
               f'{pr_ok_db}/{pr_tot_db}'   if tem_db  else '',
               f'{pr_ok_dba}/{pr_tot_dba}' if tem_dba else '',
               f'{pr_ok_rl}/{pr_tot_rl}'   if tem_rl  else '', bold=True)

    # Cabeçalhos de coluna dinâmicos
    ag_col_header  = '<th>AG</th>'        if tem_ag  else ''
    db_col_header  = '<th>DBSCAN</th>'    if tem_db  else ''
    dba_col_header = '<th>DBSCAN+AG</th>' if tem_dba else ''
    rl_col_header  = '<th>RL (DQN)</th>'  if tem_rl  else ''

    # Footer fretes
    footer_fretes  = f'<div class="footer-val"><span>Frete Original</span><strong class="red">{fmt_r(fr_real)}</strong></div>'
    if tem_ag:  footer_fretes += f'<div class="footer-val"><span>Frete AG</span><strong class="purple">{fmt_r(fr_ag)}</strong></div>'
    if tem_db:  footer_fretes += f'<div class="footer-val"><span>Frete DBSCAN</span><strong class="amber">{fmt_r(fr_db)}</strong></div>'
    if tem_dba: footer_fretes += f'<div class="footer-val"><span>Frete DBSCAN+AG</span><strong class="cyan">{fmt_r(fr_dba)}</strong></div>'
    if tem_rl:  footer_fretes += f'<div class="footer-val"><span>Frete RL (DQN)</span><strong class="teal">{fmt_r(fr_rl)}</strong></div>'

    # ── Seção Interestadual no HTML ───────────────────────────
    _df_inter_ok = ('df_interestadual' in globals() and
                    isinstance(df_interestadual, pd.DataFrame) and
                    not df_interestadual.empty)

    if _df_inter_ok:
        _inter_rows = ''
        for _, _row in df_interestadual.iterrows():
            _inter_rows += (
                f'<tr>'
                f'<td>{_row.get("Viagem","")}</td>'
                f'<td>{_row.get("Região","")}</td>'
                f'<td style="font-weight:600">{_row.get("UF","")}</td>'
                f'<td>{_row.get("Cliente","")}</td>'
                f'<td>{_row.get("Caminhão","")}</td>'
                f'<td class="num">{_row.get("Qtd Engradados",0)}</td>'
                f'<td class="num">{_row.get("Ocupação %",0):.1f}%</td>'
                f'<td class="num">{_row.get("Distancia KM (est)",0):.0f}</td>'
                f'<td class="num" style="color:#fbbf24;font-weight:600">'
                f'R$ {_row.get("Frete Estimado (R$)",0):,.2f}</td>'
                f'<td style="color:#64748b;font-size:11px">{_row.get("UFs na Viagem","")}</td>'
                f'</tr>'
            )
        _n_inter_trips = df_interestadual['Viagem'].nunique()
        _frete_inter   = df_interestadual['Frete Estimado (R$)'].sum()
        _regioes_inter = ', '.join(sorted(df_interestadual['Região'].unique()))
        _html_inter_section = f"""
<div class="table-wrap" style="margin-top:24px;border-color:#fbbf24">
  <div class="table-title" style="color:#fbbf24;border-bottom-color:#fbbf24;background:#1a1a0e">
    ✈️ Entregas Interestaduais — Transportadora Terceirizada (rodovia / CIF·FOB)
    <span style="font-size:12px;font-weight:400;margin-left:16px;color:#92400e">
      {_n_inter_trips} trip(s) &middot; Regiões: {_regioes_inter}
      &middot; Frete est.: R$ {_frete_inter:,.2f}
    </span>
  </div>
  <table>
    <thead>
      <tr>
        <th style="text-align:left">Viagem</th>
        <th style="text-align:left">Região</th>
        <th style="text-align:left">UF</th>
        <th style="text-align:left">Cliente</th>
        <th style="text-align:left">Caminhão</th>
        <th>Eng.</th>
        <th>Ocup.</th>
        <th>Dist (km)</th>
        <th>Frete Est.</th>
        <th style="text-align:left">UFs na Viagem</th>
      </tr>
    </thead>
    <tbody>{_inter_rows}</tbody>
  </table>
</div>"""
    else:
        _html_inter_section = ''


    # ── Faixa histórico DQN no rodapé (v82) ──────────────────
    # Lê rl_historico.csv para exibir progresso de aprendizado
    _rl_hist_html = ''
    if tem_rl:
        try:
            import csv as _csv_sum
            _hist_path_sum = os.path.join(DRIVE_PATH, 'rl_historico.csv')
            if os.path.exists(_hist_path_sum):
                with open(_hist_path_sum, 'r', encoding='utf-8') as _hf_sum:
                    _rows_sum = list(_csv_sum.DictReader(_hf_sum))
                if _rows_sum:
                    _last   = _rows_sum[-1]
                    _n_sess = len(_rows_sum)
                    _ep_ac  = int(_last.get('ep_acumulados', 0))
                    _eps_f  = float(_last.get('epsilon_final', 1.0))
                    _best_r = float(_last.get('melhor_reward', 0))
                    _fr_his = float(_last.get('frete_rl', 0))
                    _dt_his = _last.get('data', '') + ' ' + _last.get('hora', '')
                    _rl_hist_html = f'''
<div class="rl-hist">
  <span class="rl-hist-title">&#129302; DQN<br>Aprendizado</span>
  <div class="rl-stat">
    <span class="rl-stat-val">{_ep_ac:,}</span>
    <span class="rl-stat-lbl">Ep. acumulados</span>
  </div>
  <div class="rl-stat">
    <span class="rl-stat-val">{_n_sess}</span>
    <span class="rl-stat-lbl">Sessões rodadas</span>
  </div>
  <div class="rl-stat">
    <span class="rl-stat-val amber">{_eps_f:.4f}</span>
    <span class="rl-stat-lbl">ε final (última sessão)</span>
  </div>
  <div class="rl-stat">
    <span class="rl-stat-val slate">{_best_r:+.0f}</span>
    <span class="rl-stat-lbl">Melhor reward</span>
  </div>
  <div class="rl-stat">
    <span class="rl-stat-val">R$ {_fr_his:,.0f}</span>
    <span class="rl-stat-lbl">Frete última sessão</span>
  </div>
  <div class="rl-stat" style="border-right:none">
    <span class="rl-stat-val slate" style="font-size:13px">{_dt_his}</span>
    <span class="rl-stat-lbl">Última execução</span>
  </div>
</div>'''
        except Exception:
            _rl_hist_html = ''

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Sumário Logístico</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     background:#0f172a;color:#e2e8f0;padding:28px 32px;min-height:100vh}}
.header{{margin-bottom:24px}}
.header h1{{font-size:26px;font-weight:800;color:#1D9E75;margin-bottom:6px;
            letter-spacing:-.3px}}
.header p{{font-size:13px;color:#64748b}}
.kpi-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:28px}}
.kpi-card{{background:#1e293b;border:1px solid #334155;border-radius:14px;padding:18px 20px}}
.kpi-label{{font-size:14px;color:#e2e8f0;font-weight:700;margin-bottom:12px;
            letter-spacing:-.1px}}
.kpi-row{{display:flex;gap:8px}}
.kpi-col{{flex:1;text-align:center}}
.kpi-head{{font-size:12px;color:#1D9E75;font-weight:700;margin-bottom:5px;
           text-transform:uppercase;letter-spacing:.06em}}
.kpi-val{{font-size:22px;font-weight:700;line-height:1}}
.kpi-val.red{{color:#f87171}}
.kpi-val.green{{color:#34d399}}
.kpi-val.purple{{color:#a78bfa}}
.kpi-delta{{font-size:11px;margin-top:4px;font-weight:600}}
.kpi-delta.good{{color:#34d399}}
.kpi-delta.bad{{color:#f87171}}
.kpi-delta.neutral{{color:#64748b}}
.table-wrap{{background:#1e293b;border:1px solid #334155;border-radius:14px;
             overflow:hidden;margin-bottom:20px}}
.table-title{{padding:16px 20px;font-size:16px;font-weight:700;color:#1D9E75;
              border-bottom:2px solid #1D9E75;background:#162032;
              letter-spacing:-.1px}}
table{{width:100%;border-collapse:collapse}}
th{{padding:12px 16px;text-align:right;font-size:13px;font-weight:800;
    color:#1D9E75;text-transform:uppercase;letter-spacing:.07em;
    background:#162032;border-bottom:1px solid #334155}}
th:first-child{{text-align:left}}
td{{padding:10px 16px;font-size:13px;border-bottom:1px solid #1f2d45;
    color:#cbd5e1}}
td:first-child{{color:#f1f5f9;font-weight:500}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.pct-good{{color:#34d399;font-weight:700}}
tr:hover td{{background:rgba(29,158,117,0.06)}}
tr.tr-sub td{{color:#94a3b8;font-size:12px}}
tr.tr-sub td:first-child{{color:#94a3b8;font-weight:400}}
tr.tr-sep td{{border-bottom:1px solid #334155;padding:2px 0}}
.footer-vals{{display:flex;gap:16px;flex-wrap:wrap;margin-top:8px}}
.footer-val{{background:#1e293b;border:1px solid #334155;border-radius:12px;
             padding:14px 20px;flex:1;min-width:180px}}
.footer-val span{{display:block;font-size:12px;color:#64748b;margin-bottom:5px;
                  font-weight:600;text-transform:uppercase;letter-spacing:.05em}}
.footer-val strong{{font-size:20px;font-weight:800}}
.footer-val strong.red{{color:#f87171}}
.footer-val strong.green{{color:#34d399}}
.footer-val strong.purple{{color:#a78bfa}}
.footer-val strong.amber{{color:#fbbf24}}
.footer-val strong.cyan{{color:#22d3ee}}
.footer-val strong.teal{{color:#2dd4bf}}
.kpi-val.amber{{color:#fbbf24}}
.kpi-val.cyan{{color:#22d3ee}}
.kpi-val.teal{{color:#2dd4bf}}
.footer{{text-align:center;font-size:11px;color:#334155;margin-top:24px}}
.rl-hist{{background:#0f1e1a;border:1px solid #134e4a;border-radius:14px;
             padding:16px 24px;margin:16px 0 0;display:flex;align-items:center;
             gap:0;flex-wrap:wrap}}
.rl-hist-title{{font-size:11px;font-weight:700;color:#2dd4bf;text-transform:uppercase;
                letter-spacing:.08em;margin-right:24px;min-width:110px}}
.rl-stat{{display:flex;flex-direction:column;align-items:center;flex:1;min-width:120px;
          padding:0 12px;border-right:1px solid #134e4a}}
.rl-stat:last-child{{border-right:none}}
.rl-stat-val{{font-size:20px;font-weight:800;color:#2dd4bf;margin-bottom:2px}}
.rl-stat-val.amber{{color:#fbbf24}}
.rl-stat-val.slate{{color:#94a3b8}}
.rl-stat-lbl{{font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:.05em;text-align:center}}
</style>
</head>
<body>

<div class="header">
  <h1>📊 Sumário — Comparativo de Planos Logísticos</h1>
  <p>Sistema de Otimização Logística &middot; {hoje}</p>
</div>

<div class="kpi-grid">
{kpi_html}
</div>

<div class="table-wrap">
  <div class="table-title">Detalhamento por métrica e tipo de caminhão</div>
  <table>
    <thead>
      <tr>
        <th style="text-align:left">Métrica</th>
        <th>Original</th>
        {ag_col_header}
        {db_col_header}
        {dba_col_header}
        {rl_col_header}
      </tr>
    </thead>
    <tbody>
      {rows}
    </tbody>
  </table>
</div>

<div class="footer-vals">
  {footer_fretes}
</div>

{_html_inter_section}

{_rl_hist_html}

<div class="footer">
  Sistema de Otimização Logística &middot; Gerado em {hoje}
</div>

</body>
</html>"""

    with open(caminho_saida, 'w', encoding='utf-8') as f:
        f.write(html)
    salvar_e_abrir(caminho_saida, silent=True)
    print(f"   📊 Sumário HTML salvo → {os.path.basename(caminho_saida)}")


# =============================================================
# SEÇÃO 13F — PLANO RL (Deep Q-Network — Estado Fixo)
# Agente aprende a construir rotas por reforço — comparação com AG
#
# Arquitetura DQN com estado fixo de 12 features:
#   Estado  : SEMPRE 12 features — independente do N de destinos do dia
#             [lat_atual, lon_atual,          ← posição do agente
#              lat_melhor, lon_melhor,         ← destino mais urgente
#              urg_melhor, vol_melhor, dist_melhor,
#              n_restantes, urg_media, vol_total, dist_media,
#              hora_atual]
#   Ação    : índice do próximo destino (0..N-1), inválidos mascarados
#   Reward  : -(frete_perna) - penalidade_atraso - co2_perna × 0.1
#   Backbone: Linear(12→256→128) — salvo e reutilizado entre dias
#   Head    : Linear(128→N) — recriado a cada execução (N muda)
#   Benefício: modelo acumula aprendizado com NFs de qualquer data
# =============================================================
print("\n🤖 Plano RL (DQN) — treinando agente de roteização...")

df_rl          = pd.DataFrame()
timeline_rl    = []
co2_rl_kg      = 0.0
frete_rl       = 0.0

try:
    import torch
    import torch.nn as _nn
    import torch.optim as _optim
    import collections, random as _rnd_rl, math as _math_rl

    # ── Grupos de destino (mesma base dos outros planos) ─────
    _rl_grupos = construir_grupos_destino(eng_final)
    for _g in _rl_grupos:
        _g['limite'] = limite_otimizado_cliente(_g['cli'])
    _rl_n = len(_rl_grupos)

    if _rl_n < 2:
        print("   ⚠️  Poucos destinos para DQN — pulando.")
    else:
        _rl_seed = 7
        torch.manual_seed(_rl_seed)
        _rnd_rl.seed(_rl_seed)

        # ── Estado fixo de 12 features — independente de N ───
        #
        # O estado NUNCA depende do número de destinos do dia.
        # Isso permite que o modelo acumule aprendizado entre
        # execuções com qualquer quantidade de NFs.
        #
        # Features:
        #  0-1  : posição atual do agente (lat, lon normalizados)
        #  2-3  : destino mais urgente disponível (lat, lon)
        #  4    : urgência do destino mais urgente (norm)
        #  5    : volume do destino mais urgente (norm)
        #  6    : distância do agente ao destino mais urgente (norm)
        #  7    : n° de destinos restantes (norm 0-1)
        #  8    : urgência MÉDIA dos restantes (norm)
        #  9    : volume TOTAL restante (norm)
        #  10   : distância MÉDIA dos restantes ao armazém (norm)
        #  11   : hora atual (norm 0-1 dentro do dia útil)
        #
        _STATE_SIZE  = 12          # SEMPRE 12 — não muda com N
        _ACTION_SIZE = _rl_n       # ações = destinos do dia (mascara inválidos)

        _arm_lat  = ARMAZEM_SUZANO['lat']
        _arm_lon  = ARMAZEM_SUZANO['lon']
        _vols_rl  = [sum(e['vol'] for e in g['engs']) for g in _rl_grupos]
        _lims_rl  = [lim_min(g['limite']) for g in _rl_grupos]
        _vol_max  = max(_vols_rl) if _vols_rl else 1.0
        _vol_tot  = sum(_vols_rl)
        _hora_ini = HORA_INICIO * 60
        _hora_fim = 26 * 60   # 02:00 do dia seguinte

        def _norm_lat(v):  return (v + 90)  / 180
        def _norm_lon(v):  return (v + 180) / 360
        def _norm_hor(v):  return max(0.0, min(1.0, (v - _hora_ini) / (_hora_fim - _hora_ini + 1e-9)))
        def _norm_vol(v):  return v / (_vol_max + 1e-9)
        def _norm_dist(v): return min(1.0, v / 2000)   # normaliza até 2000 km

        def _build_state(visitados, hora_min, pos_lat, pos_lon):
            """
            Constrói vetor de estado fixo de 12 features.
            Funciona para qualquer N de destinos — modelo sempre compatível.
            """
            disponiveis = [i for i in range(_rl_n) if i not in visitados]

            if not disponiveis:
                return torch.zeros(_STATE_SIZE, dtype=torch.float32)

            # Destino mais urgente disponível
            _melhor_idx = min(disponiveis, key=lambda i: _lims_rl[i])
            _gm = _rl_grupos[_melhor_idx]
            _dm, _ = _haversine((pos_lat, pos_lon), (_gm['lat'], _gm['lon']))

            # Agregados dos restantes
            _n_rest     = len(disponiveis)
            _urg_media  = sum(_lims_rl[i] for i in disponiveis) / _n_rest
            _vol_rest   = sum(_vols_rl[i] for i in disponiveis)
            _dist_media = sum(
                _haversine((_arm_lat, _arm_lon), (_rl_grupos[i]['lat'], _rl_grupos[i]['lon']))[0]
                for i in disponiveis
            ) / _n_rest

            return torch.tensor([
                _norm_lat(pos_lat),
                _norm_lon(pos_lon),
                _norm_lat(_gm['lat']),
                _norm_lon(_gm['lon']),
                _norm_hor(_lims_rl[_melhor_idx]),
                _norm_vol(_vols_rl[_melhor_idx]),
                _norm_dist(_dm),
                _n_rest / max(_rl_n, 1),
                _norm_hor(_urg_media),
                _norm_vol(_vol_rest / max(_rl_n, 1)),
                _norm_dist(_dist_media),
                _norm_hor(hora_min),
            ], dtype=torch.float32)

        # ── Rede DQN ─────────────────────────────────────────
        # Entrada: 12 features fixas
        # Saída:   N ações (destinos do dia) — varia por execução
        # A camada de saída é a ÚNICA que muda com N.
        # Os pesos das camadas ocultas (12→256→128) são sempre
        # compatíveis entre execuções — é onde fica o aprendizado.
        class _DQN(_nn.Module):
            def __init__(self, n_in, n_out):
                super().__init__()
                self.backbone = _nn.Sequential(
                    _nn.Linear(n_in, 256), _nn.ReLU(),
                    _nn.Linear(256, 128),  _nn.ReLU(),
                )
                self.head = _nn.Linear(128, n_out)
            def forward(self, x):
                return self.head(self.backbone(x))

        _policy_net = _DQN(_STATE_SIZE, _ACTION_SIZE)
        _target_net = _DQN(_STATE_SIZE, _ACTION_SIZE)
        _optimizer  = _optim.Adam(_policy_net.parameters(), lr=RL_LR)
        _criterion  = _nn.MSELoss()
        _buffer     = collections.deque(maxlen=RL_BUFFER_MAX)

        # ── Persistência de pesos ─────────────────────────────
        # STATE_SIZE é sempre 12 — modelo sempre compatível entre dias.
        # Só o head (128→N) muda; o backbone (12→256→128) é carregado.
        _rl_model_path    = os.path.join(DRIVE_PATH, 'RL_DQN_model.pt')
        _rl_episodio_base = 0

        if os.path.exists(_rl_model_path):
            try:
                import numpy as _np_rl
                import torch.serialization as _ts_rl
                _ts_rl.add_safe_globals([getattr(getattr(_np_rl, "_core", getattr(_np_rl, "core", None)), "multiarray", object).__dict__.get("scalar", object)])
                try:
                    _checkpoint = torch.load(_rl_model_path, weights_only=True)
                except Exception:
                    _checkpoint = torch.load(_rl_model_path, weights_only=False)
                _ck_state   = _checkpoint.get('state_size', 0)
                if _ck_state == _STATE_SIZE:
                    # Carrega só o backbone — head é recriado para N do dia
                    _bb_state = {
                        k: v for k, v in _checkpoint['policy_state'].items()
                        if k.startswith('backbone.')
                    }
                    _missing, _unexpected = _policy_net.load_state_dict(
                        _bb_state, strict=False)
                    _rl_episodio_base = _checkpoint.get('episodios_total', 0)
                    _ep_acum          = _rl_episodio_base
                    print(f"   📂 Backbone RL carregado → {_ep_acum} ep acumulados")
                    print(f"      Head adaptado para {_rl_n} destinos do dia")
                    print(f"      Continuando treino por mais {RL_EPISODIOS} episódios...")
                else:
                    # state_size 12 deve sempre bater — só falha se arquivo for de versão antiga
                    print(f"   ⚠️  Modelo legado (state_size={_ck_state}) — removendo e reiniciando")
                    try:
                        os.remove(_rl_model_path)
                        print(f"      RL_DQN_model.pt removido — próxima execução começa limpa")
                    except Exception:
                        pass
                    print(f"      Treinando do zero com nova arquitetura...")
            except Exception as _e_load:
                print(f"   ⚠️  Erro ao carregar modelo: {type(_e_load).__name__} — treinando do zero")
                try:
                    os.remove(_rl_model_path)
                    print(f"      Arquivo corrompido removido — próxima execução começa limpa")
                except Exception:
                    pass
        else:
            print(f"   🆕 Nenhum modelo salvo — treinando do zero (estado fixo 12 features)")

        _target_net.load_state_dict(_policy_net.state_dict())
        _target_net.eval()

        # ── Simulação de um episódio ──────────────────────────
        def _simular_episodio(politica='greedy'):
            """
            Constrói uma rota completa seguindo a política do agente.
            politica='greedy' usa argmax(Q); 'epsilon' usa ε-greedy.
            Retorna (sequencia_indices, recompensa_total, chegadas_dict).
            """
            visitados  = set()
            sequencia  = []
            hora_min   = HORA_INICIO * 60
            pos        = (_arm_lat, _arm_lon)
            reward_tot = 0.0
            chegadas   = {}

            while len(visitados) < _rl_n:
                estado = _build_state(visitados, hora_min, pos[0], pos[1])

                # Máscara de ações válidas
                mascara = torch.full((_ACTION_SIZE,), float('-inf'))
                disponiveis = [i for i in range(_rl_n) if i not in visitados]
                for i in disponiveis:
                    mascara[i] = 0.0

                if politica == 'epsilon' and _rnd_rl.random() < _epsilon:
                    acao = _rnd_rl.choice(disponiveis)
                else:
                    with torch.no_grad():
                        q = _policy_net(estado) + mascara
                    acao = int(q.argmax().item())

                # Transição de estado
                dest_g  = _rl_grupos[acao]
                dest_pos = (dest_g['lat'], dest_g['lon'])
                km_perna, mins_perna = _haversine(pos, dest_pos)

                chegada_bruta = datetime.today().replace(
                    hour=0, minute=0, second=0, microsecond=0
                ) + timedelta(minutes=hora_min + mins_perna)
                chegada_dt = ajustar_pausas(
                    datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
                    + timedelta(minutes=hora_min),
                    chegada_bruta
                )
                hora_chegada = dt_min(chegada_dt)
                no_prazo     = hora_chegada <= lim_min(dest_g['limite'])

                # Frete desta perna
                _truck_est, _, _ = empacotar_em_caminhao([dest_g], truck_list)
                tipo_est  = _truck_est['tipo'] if _truck_est else truck_list[-1]['tipo']
                frete_perna = calcular_frete_viagem(km_perna, tipo_est, tabela_frete) \
                              if tabela_frete else km_perna * 3.5

                # CO2 desta perna
                consumo_tipo = consumo_combustivel.get(tipo_est, 0)
                co2_perna    = km_perna * consumo_tipo * CO2_KG_POR_LITRO

                # Recompensa normalizada por N — evita valores muito negativos
                # que dificultam o aprendizado da rede com muitos destinos
                reward = -(frete_perna + co2_perna * 0.1) / max(_rl_n, 1)
                if not no_prazo:
                    reward -= RL_PENALIDADE_ATR / max(_rl_n, 1)

                visitados.add(acao)
                sequencia.append(acao)
                chegadas[acao] = hora_chegada
                hora_min = hora_chegada + dest_g.get('descarga', TEMPO_DESCARGA_MIN)
                pos = dest_pos

                # v81: inclui custo de retorno ao armazém no último passo
                # O AG/DBSCAN+AG otimizam a rota completa; sem isso o DQN
                # subestima rotas que terminam longe do armazém.
                done = (len(visitados) == _rl_n)
                if done:
                    _km_volta, _ = _haversine(pos, (_arm_lat, _arm_lon))
                    _trk_volta, _, _ = empacotar_em_caminhao([dest_g], truck_list)
                    _tipo_volta = _trk_volta['tipo'] if _trk_volta else truck_list[-1]['tipo']
                    _frete_volta = (calcular_frete_viagem(_km_volta, _tipo_volta, tabela_frete)
                                    if tabela_frete else _km_volta * 3.5)
                    reward -= _frete_volta / max(_rl_n, 1)

                reward_tot += reward

                # Armazena transição no replay buffer (durante treino)
                if politica == 'epsilon':
                    prox_vis = set(visitados)
                    prox_estado = _build_state(prox_vis, hora_min, pos[0], pos[1])
                    _buffer.append((estado, acao, reward, prox_estado, done))

            return sequencia, reward_tot, chegadas

        # ── Loop de treino ────────────────────────────────────
        # v86: epsilon sempre resetado para 1.0 a cada nova sessão
        # O backbone (pesos) é preservado entre sessões,
        # mas a exploração recomeça do zero para aproveitar os 5.000
        # episódios integralmente — sem isso o agente roda greedy puro.
        _epsilon = RL_EPSILON_INI   # sempre 1.0, ignora valor salvo no .pt
        if _rl_episodio_base > 0:
            print(f"   ε resetado para {_epsilon:.3f} (backbone de {_rl_episodio_base} ep carregado — exploração reinicia)")

        _historico_rl = []
        _melhor_seq   = None
        _melhor_rw    = float('-inf')

        # ── Warm-start: pré-popula buffer com sequência DBSCAN+AG ──
        # O agente começa aprendendo a partir de uma boa solução
        # em vez de exploração puramente aleatória.
        _ws_feito = False
        _ws_seq   = None

        # Tenta usar sequência do DBSCAN+AG (melhor plano conhecido)
        if '_melhor2_ind' in globals() and _melhor2_ind:
            # Mapeia índices do DBSCAN+AG para índices do RL
            # (ambos usam construir_grupos_destino sobre eng_final)
            _ws_seq = _melhor2_ind[:]
            _ws_fonte = 'DBSCAN+AG'
        elif 'melhor_ind' in globals() and melhor_ind:
            _ws_seq = melhor_ind[:]
            _ws_fonte = 'AG'

        if _ws_seq and len(_ws_seq) == _rl_n:
            print(f"   🔥 Warm-start com sequência {_ws_fonte} "
                  f"({RL_BUFFER_MAX // _rl_n} replay(s) × {_rl_n} passos)...")
            _ws_rw_total = 0.0
            # Popula o buffer com N repetições da sequência seed
            # com pequenas variações (mutações leves) para diversidade
            # v81: limitado a ~20% do buffer para evitar que o agente
            # convirja para cópia ruidosa do DBSCAN+AG (era: RL_BUFFER_MAX // N)
            _n_replays = max(2, (RL_BUFFER_MAX // max(_rl_n, 1)) // 5)
            for _ws_rep in range(_n_replays):
                # Varia levemente a sequência (2-opt aleatório)
                _seq_ws = _ws_seq[:]
                if _ws_rep > 0:  # primeira replica é pura
                    _n_swaps = max(1, _rl_n // 6)
                    for _ in range(_n_swaps):
                        _a, _b = _rnd_rl.sample(range(_rl_n), 2)
                        _seq_ws[_a], _seq_ws[_b] = _seq_ws[_b], _seq_ws[_a]

                # Simula a sequência e popula o buffer
                _ws_vis  = set()
                _ws_hora = HORA_INICIO * 60
                _ws_pos  = (_arm_lat, _arm_lon)
                _ws_rw   = 0.0

                for _ws_idx in _seq_ws:
                    if _ws_idx >= _rl_n:
                        continue
                    _ws_estado = _build_state(_ws_vis, _ws_hora, _ws_pos[0], _ws_pos[1])
                    _ws_g      = _rl_grupos[_ws_idx]
                    _ws_dest   = (_ws_g['lat'], _ws_g['lon'])
                    _ws_km, _ws_mins = _haversine(_ws_pos, _ws_dest)

                    _ws_cheg_b = datetime.today().replace(
                        hour=0, minute=0, second=0, microsecond=0
                    ) + timedelta(minutes=_ws_hora + _ws_mins)
                    _ws_cheg = ajustar_pausas(
                        datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
                        + timedelta(minutes=_ws_hora), _ws_cheg_b)
                    _ws_no_prazo = dt_min(_ws_cheg) <= lim_min(_ws_g['limite'])

                    _ws_trk, _, _ = empacotar_em_caminhao([_ws_g], truck_list)
                    _ws_tipo = _ws_trk['tipo'] if _ws_trk else truck_list[-1]['tipo']
                    _ws_frete = (calcular_frete_viagem(_ws_km, _ws_tipo, tabela_frete)
                                 if tabela_frete else _ws_km * 3.5)
                    _ws_co2   = _ws_km * consumo_combustivel.get(_ws_tipo, 0) * CO2_KG_POR_LITRO
                    _ws_r     = -(_ws_frete + _ws_co2 * 0.1) / max(_rl_n, 1)
                    if not _ws_no_prazo:
                        _ws_r -= RL_PENALIDADE_ATR / max(_rl_n, 1)

                    _ws_vis.add(_ws_idx)
                    _ws_hora = dt_min(_ws_cheg) + _ws_g.get('descarga', TEMPO_DESCARGA_MIN)
                    _ws_pos  = _ws_dest
                    _ws_rw  += _ws_r

                    _ws_prox_estado = _build_state(_ws_vis, _ws_hora, _ws_pos[0], _ws_pos[1])
                    _ws_done        = (len(_ws_vis) == _rl_n)
                    _buffer.append((_ws_estado, _ws_idx, _ws_r, _ws_prox_estado, _ws_done))

                _ws_rw_total += _ws_rw

            _ws_rw_media = _ws_rw_total / max(_n_replays, 1)
            _melhor_seq  = _ws_seq[:]
            _melhor_rw   = _ws_rw_media
            print(f"   ✅ Buffer pré-populado: {len(_buffer)} transições "
                  f"| reward seed: {_ws_rw_media:+.0f}")
            _ws_feito = True
        else:
            print("   ℹ️  Warm-start indisponível — DBSCAN+AG ainda não rodou "
                  "ou N de destinos diverge. Iniciando exploração aleatória.")

        print(f"   Parâmetros: ep={RL_EPISODIOS} | ε {_epsilon:.3f}→{RL_EPSILON_MIN}"
              f" | γ={RL_GAMMA} | lr={RL_LR} | batch={RL_BATCH}"
              f"{' | warm-start ✅' if _ws_feito else ' | sem warm-start'}")

        for _ep in range(1, RL_EPISODIOS + 1):
            seq, rw, _ = _simular_episodio('epsilon')

            if rw > _melhor_rw:
                _melhor_rw  = rw
                _melhor_seq = seq[:]

            # Treino com replay buffer
            if len(_buffer) >= RL_BATCH:
                _batch = _rnd_rl.sample(_buffer, RL_BATCH)
                _s  = torch.stack([b[0] for b in _batch])
                _a  = torch.tensor([b[1] for b in _batch], dtype=torch.long)
                _r  = torch.tensor([b[2] for b in _batch], dtype=torch.float32)
                _ns = torch.stack([b[3] for b in _batch])
                _d  = torch.tensor([b[4] for b in _batch], dtype=torch.float32)

                _q_curr = _policy_net(_s).gather(1, _a.unsqueeze(1)).squeeze(1)
                with torch.no_grad():
                    _q_next = _target_net(_ns).max(1)[0]
                _q_tgt = _r + RL_GAMMA * _q_next * (1 - _d)

                _loss = _criterion(_q_curr, _q_tgt)
                _optimizer.zero_grad()
                _loss.backward()
                torch.nn.utils.clip_grad_norm_(_policy_net.parameters(), 1.0)  # v81: gradient clipping
                _optimizer.step()

            # Atualiza rede target
            if _ep % RL_TARGET_UPDATE == 0:
                _target_net.load_state_dict(_policy_net.state_dict())

            # Decaimento epsilon
            _epsilon = max(RL_EPSILON_MIN, _epsilon * RL_EPSILON_DECAY)

            _historico_rl.append((_ep, rw, _melhor_rw))

            if _ep % 100 == 0 or _ep == RL_EPISODIOS:
                print(f"   Ep {_ep:4d}/{RL_EPISODIOS} | ε={_epsilon:.3f}"
                      f" | reward={rw:+.0f} | melhor={_melhor_rw:+.0f}")

        print(f"\n   🏆 Melhor sequência RL encontrada após {RL_EPISODIOS} episódios")

        # ── Salva pesos para a próxima execução (v84) ───────
        # Salva o backbone completo (12→256→128) — sempre compatível.
        # O head (128→N) não é salvo pois N muda entre dias.
        # v84: salvamento fora do try/except interno para garantir execução
        #      e expor erros reais ao invés de suprimi-los silenciosamente.
        _episodios_total = _rl_episodio_base + RL_EPISODIOS
        _save_ok = False
        try:
            _save_dict = {
                'policy_state':    _policy_net.state_dict(),
                'optimizer_state': _optimizer.state_dict(),
                'episodios_total': _episodios_total,
                'state_size':      _STATE_SIZE,
                'melhor_reward':   _melhor_rw,
                'epsilon':         _epsilon,
                'n_destinos_dia':  _rl_n,
            }
            torch.save(_save_dict, _rl_model_path)
            _save_ok = True
            print(f"   💾 Backbone salvo → RL_DQN_model.pt "
                  f"({_episodios_total} ep acumulados | N={_rl_n} destinos hoje | ε={_epsilon:.3f})")
        except Exception as _e_save:
            print(f"   ❌ ERRO ao salvar modelo: {repr(_e_save)}")
            import traceback as _tb_save
            _tb_save.print_exc()
        finally:
            if not _save_ok:
                print(f"   ⚠️  Modelo NÃO salvo — verifique permissões em {_rl_model_path}")
        # ── Decodifica melhor sequência em timeline ───────────
        _rl_h_ini    = datetime.today().replace(
            hour=HORA_INICIO, minute=0, second=0, microsecond=0)
        _rl_prox_doc = _rl_h_ini
        _rl_prox_pat = _rl_h_ini
        id_trip_rl   = 1

        # Agrupa sequência em trips (respeita capacidade do caminhão)
        _rl_trips = []
        _i = 0
        while _i < len(_melhor_seq):
            _trip_rl = []
            for _j in range(_i, len(_melhor_seq)):
                _trip_rl.append(_rl_grupos[_melhor_seq[_j]])
                _trk_test, _aloc_test, _ = empacotar_em_caminhao(_trip_rl, truck_list)
                if not _aloc_test:
                    _trip_rl = _trip_rl[:-1]
                    break
                if len(_trip_rl) >= AG_MAX_PARADAS:
                    break
            if not _trip_rl:
                _trip_rl = [_rl_grupos[_melhor_seq[_i]]]
            _rl_trips.append(_trip_rl)
            _i += len(_trip_rl)

        for _trip_rl in _rl_trips:
            _trk_rl, _aloc_rl, _ = empacotar_em_caminhao(_trip_rl, truck_list)
            if not _trk_rl or not _aloc_rl:
                continue

            _qtd_e_rl = sum(len(g['engs']) for g in _trip_rl)
            _tc_rl    = tempo_carregamento(_qtd_e_rl)
            _loc_rl   = _prox_carga(_trk_rl['tipo'])
            if _loc_rl == 'Pateo':
                _hi_rl       = _rl_prox_pat
                _hs_rl       = _hi_rl + timedelta(minutes=_tc_rl)
                _rl_prox_pat = _hs_rl
            else:
                _hi_rl       = _rl_prox_doc
                _hs_rl       = _hi_rl + timedelta(minutes=_tc_rl)
                _rl_prox_doc = _hs_rl

            _rot_rl, _, _ = simular_viagem(
                _trip_rl, mem_cache, gmaps, API_DISPONIVEL, controlador_api,
                h_inicio=_hs_rl)

            _vol_rl = min(
                sum((e['dim_f'][0]*e['dim_f'][1]*e['dim_f'][2])
                    if 'dim_f' in e else e.get('vol', 0) for e in _aloc_rl),
                _trk_rl['vol'])

            _cm_rl  = {}
            for _e in _aloc_rl:
                _cm_rl.setdefault(_e['cli'], []).append(_e)

            _pr_rl  = (ARMAZEM_SUZANO['lat'], ARMAZEM_SUZANO['lon'])
            _hr_rl  = _hs_rl
            _tid_rl = f"RL_{id_trip_rl:03d}"

            for _g in _trip_rl:
                _cli   = _g['cli']
                _engs  = _cm_rl.get(_cli, [])
                if not _engs:
                    continue
                _dst_rl  = (_g['lat'], _g['lon'])
                _rid_rl  = _rota_id(_pr_rl, _dst_rl)
                _km_rl, _m_rl = _rot_rl.get(_rid_rl, _haversine(_pr_rl, _dst_rl))
                _cb_rl   = _hr_rl + timedelta(minutes=_m_rl)
                _ca_rl   = ajustar_pausas(_hr_rl, _cb_rl)
                _ok_rl   = dt_min(_ca_rl) <= lim_min(_g['limite'])

                timeline_rl.append({
                    'Viagem':              _tid_rl,
                    'Caminhão':            _trk_rl['tipo'],
                    'Vol Caminhão (m³)':   round(_trk_rl['vol'], 2),
                    'Vol Carga Total (m³)':round(_vol_rl, 4),
                    'Ocupação %':          round(_vol_rl / _trk_rl['vol'] * 100, 1),
                    'Cliente':             _cli,
                    'Qtd Engradados':      len(_engs),
                    'Tipos Engradados':    ", ".join(sorted(set(e['tipo'] for e in _engs))),
                    'NFs':                 "/".join(sorted(set(e['nf'] for e in _engs))),
                    'Início Carga':        _hi_rl.strftime("%H:%M"),
                    'Local Carga':         _loc_rl,
                    'Tempo Carga (min)':   _tc_rl,
                    'Saída Armazém':       _hs_rl.strftime("%H:%M"),
                    'Distancia KM':        round(_km_rl, 2),
                    'Chegada':             _ca_rl.strftime("%H:%M"),
                    'Tempo Descarga (min)': _g.get('descarga', TEMPO_DESCARGA_MIN),
                    'Limite Entrega':      _g['limite'],
                    'Dentro do Prazo':     '✅' if _ok_rl else '❌',
                    'Fonte Rota':          'Cache/Haversine'
                })

                _hr_rl = _ca_rl + timedelta(minutes=_g.get('descarga', TEMPO_DESCARGA_MIN))
                _pr_rl = _dst_rl

            id_trip_rl += 1

        df_rl = pd.DataFrame(timeline_rl)
        if not df_rl.empty:
            df_rl, co2_rl_kg, _co2_rl_t = calcular_co2(df_rl, consumo_combustivel)
            df_rl, frete_rl = calcular_frete_total(df_rl, tabela_frete)
            salvar_excel(df_rl,
                         os.path.join(DRIVE_PATH, '06_TIMELINE_RL.xlsx'),
                         '06_TIMELINE_RL.xlsx')
            _nv_rl  = df_rl['Viagem'].nunique()
            _ok_rl2 = int((df_rl['Dentro do Prazo'] == '✅').sum())
            _oc_rl  = df_rl['Ocupação %'].mean()
            _km_rl2 = df_rl['Distancia KM'].sum()
            print(f"   ✅ {_nv_rl} viagem(ns) RL → 06_TIMELINE_RL.xlsx")
            print(f"   👥 Clientes/viagem : {df_rl.groupby('Viagem')['Cliente'].count().mean():.1f}")
            print(f"   ⏱️  No prazo        : {_ok_rl2}/{len(df_rl)}")
            print(f"   🚛 Ocupação média  : {_oc_rl:.1f}%")
            print(f"   📏 Distância total : {_km_rl2:.1f} km")
            print(f"   💰 Frete           : R$ {frete_rl:,.2f}")
            if co2_rl_kg > 0:
                print(f"   🌿 CO2             : {co2_rl_kg:.1f} kg")

        # ── Grava histórico de aprendizado em CSV (v83) ──────────────────
        # Movido para APÓS cálculo de frete/CO₂ — corrige bug do frete=0 (v82)
        # Arquivo: C:\logistica\rl_historico.csv — uma linha por execução
        try:
            import csv as _csv_rl
            _hist_path   = os.path.join(DRIVE_PATH, 'rl_historico.csv')
            _hist_exists = os.path.exists(_hist_path)
            _hist_frete  = frete_rl if 'frete_rl' in dir() else globals().get('frete_rl', 0.0)
            _hist_co2    = co2_rl_kg if 'co2_rl_kg' in dir() else globals().get('co2_rl_kg', 0.0)
            _hist_dist   = round(df_rl['Distancia KM'].sum(), 1) if not df_rl.empty and 'Distancia KM' in df_rl.columns else 0.0
            with open(_hist_path, 'a', newline='', encoding='utf-8') as _hf:
                _hw = _csv_rl.writer(_hf)
                if not _hist_exists:
                    _hw.writerow([
                        'data', 'hora', 'n_destinos',
                        'ep_sessao', 'ep_acumulados',
                        'melhor_reward', 'epsilon_final',
                        'frete_rl', 'distancia_km', 'co2_kg'
                    ])
                _hw.writerow([
                    datetime.now().strftime('%d/%m/%Y'),
                    datetime.now().strftime('%H:%M'),
                    _rl_n,
                    RL_EPISODIOS,
                    _episodios_total,
                    round(_melhor_rw, 2),
                    round(_epsilon, 4),
                    round(_hist_frete, 2),
                    _hist_dist,
                    round(_hist_co2, 1),
                ])
            print(f"   📈 Histórico gravado → rl_historico.csv ({_episodios_total} ep acum. | sessão #{_episodios_total // RL_EPISODIOS})")
        except Exception as _e_hist:
            print(f"   ⚠️  Erro ao gravar histórico: {repr(_e_hist)}")

        # ── Gráfico de convergência RL ────────────────────────
        try:
            import json as _json_rl
            _eps_rl   = [h[0] for h in _historico_rl]
            _rws_rl   = [h[1] for h in _historico_rl]
            _best_rl  = [h[2] for h in _historico_rl]
            _ep_ini   = _historico_rl[0][1]
            _ep_fin   = _melhor_rw
            _melhoria = abs((_ep_fin - _ep_ini) / (_ep_ini + 1e-9)) * 100
            _hoje_rl  = datetime.now().strftime('%d/%m/%Y %H:%M')

            _html_rl = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>RL DQN — Convergência</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     background:#0f172a;color:#e2e8f0;padding:28px 32px}}
.header h1{{font-size:20px;font-weight:700;color:#f1f5f9;margin-bottom:4px}}
.header p{{font-size:13px;color:#64748b;margin-bottom:20px}}
.hero{{background:linear-gradient(135deg,#0891b2 0%,#0e7490 100%);
      border-radius:14px;padding:20px 28px;margin-bottom:20px;
      display:flex;gap:40px;align-items:center;flex-wrap:wrap}}
.hero-val{{font-size:32px;font-weight:800;color:white;line-height:1}}
.hero-val.red{{color:#fca5a5}}
.hero-sub{{font-size:11px;color:rgba(255,255,255,0.7);margin-top:4px;text-transform:uppercase}}
.hero-div{{width:1px;height:50px;background:rgba(255,255,255,0.2)}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:24px}}
.card{{background:#1e293b;border-radius:12px;padding:14px 16px;border:1px solid #334155}}
.card-val{{font-size:20px;font-weight:700;color:#22d3ee;margin-bottom:3px}}
.card-lbl{{font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:.05em}}
.chart-wrap{{background:#1e293b;border-radius:14px;padding:24px;border:1px solid #334155}}
.chart-title{{font-size:14px;font-weight:600;color:#94a3b8;margin-bottom:16px}}
canvas{{max-height:360px}}
.footer{{text-align:center;font-size:11px;color:#334155;margin-top:12px}}
</style></head><body>
<div class="header">
  <h1>DQN — Aprendizado por Reforço</h1>
  <p>TTB Logistics &middot; {_hoje_rl} &middot; {_rl_n} destinos &middot; {RL_EPISODIOS} episódios</p>
</div>
<div class="hero">
  <div><div class="hero-val red">{_ep_ini:+.0f}</div><div class="hero-sub">Reward inicial</div></div>
  <div class="hero-div"></div>
  <div><div class="hero-val">{_ep_fin:+.0f}</div><div class="hero-sub">Melhor reward</div></div>
  <div class="hero-div"></div>
  <div><div class="hero-val">{_melhoria:.1f}%</div><div class="hero-sub">Melhoria obtida</div></div>
</div>
<div class="cards">
  <div class="card"><div class="card-val">{RL_EPISODIOS}</div><div class="card-lbl">Episódios</div></div>
  <div class="card"><div class="card-val">{RL_BATCH}</div><div class="card-lbl">Batch size</div></div>
  <div class="card"><div class="card-val">{RL_EPSILON_INI}→{RL_EPSILON_MIN}</div><div class="card-lbl">ε decay</div></div>
  <div class="card"><div class="card-val">{RL_GAMMA}</div><div class="card-lbl">Gamma (γ)</div></div>
</div>
<div class="chart-wrap">
  <div class="chart-title">Evolução da recompensa por episódio</div>
  <canvas id="cv"></canvas>
</div>
<div class="footer">DQN · Hidden 256→128 · Adam lr={RL_LR} · Target update={RL_TARGET_UPDATE} ep</div>
<script>
new Chart(document.getElementById('cv'),{{
  type:'line',
  data:{{
    labels:{_json_rl.dumps(_eps_rl)},
    datasets:[
      {{label:'Reward episódio',data:{_json_rl.dumps([round(v,1) for v in _rws_rl])},
        borderColor:'#38bdf8',backgroundColor:'rgba(56,189,248,0.05)',
        borderWidth:1,pointRadius:0,tension:0.2,fill:true,order:2}},
      {{label:'Melhor reward',data:{_json_rl.dumps([round(v,1) for v in _best_rl])},
        borderColor:'#22d3ee',backgroundColor:'rgba(34,211,238,0.08)',
        borderWidth:2.5,pointRadius:0,tension:0.35,fill:true,order:1}}
    ]
  }},
  options:{{
    responsive:true,
    interaction:{{mode:'index',intersect:false}},
    plugins:{{
      legend:{{labels:{{color:'#94a3b8',font:{{size:12}}}}}},
      tooltip:{{backgroundColor:'#0f172a',borderColor:'#0891b2',borderWidth:1,
               titleColor:'#f1f5f9',bodyColor:'#94a3b8',padding:10}}
    }},
    scales:{{
      x:{{ticks:{{color:'#475569',maxTicksLimit:10}},grid:{{color:'rgba(255,255,255,0.03)'}}}},
      y:{{ticks:{{color:'#475569'}},grid:{{color:'rgba(255,255,255,0.05)'}}}}
    }}
  }}
}});
</script></body></html>"""

            _path_rl_conv = os.path.join(DRIVE_PATH, 'RL_DQN_convergencia.html')
            with open(_path_rl_conv, 'w', encoding='utf-8') as _f:
                _f.write(_html_rl)
            salvar_e_abrir(_path_rl_conv, silent=True)
            print(f"   📈 Gráfico RL → RL_DQN_convergencia.html")
        except Exception as _e_rl_graf:
            print(f"   ⚠️  Erro gráfico RL: {repr(_e_rl_graf)}")

except ImportError:
    print("   ⚠️  PyTorch não instalado — DQN ignorado.")
    print("      Execute: pip install torch")
except Exception as _e_rl:
    print(f"   ⚠️  Erro no plano RL: {repr(_e_rl)}")

# Gera o sumário ao final do processamento
try:
    gerar_html_sumario(os.path.join(DRIVE_PATH, 'Sumario_Comparativo.html'))
except Exception as e_sum:
    print(f"   ⚠️  Erro ao gerar sumário HTML: {repr(e_sum)}")

controlador_api.resumo()
print("\n🏁 Processamento concluído!")
