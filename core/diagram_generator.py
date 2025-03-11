import os
import re
import pandas as pd
from graphviz import Digraph

def extrair_tabela_coluna(texto):
    match = re.match(r"'?([^']+)'?\[([^\]]+)]", texto.strip())
    return match.groups() if match else (None, None)

def gerar_diagrama(path_csv, nome_base, pasta_documentacao):
    df = pd.read_csv(path_csv, sep=';')
    df['tabela_destino'] = df['to'].apply(lambda x: extrair_tabela_coluna(x)[0])

    todas_tabelas = pd.concat([
        df['from'].apply(lambda x: extrair_tabela_coluna(x)[0]),
        df['to'].apply(lambda x: extrair_tabela_coluna(x)[0])
    ])
    tabela_central = todas_tabelas.value_counts().idxmax()

    from collections import defaultdict, deque

    grafo_conexo = defaultdict(set)
    for _, row in df.iterrows():
        t1, _ = extrair_tabela_coluna(row['from'])
        t2, _ = extrair_tabela_coluna(row['to'])
        grafo_conexo[t1].add(t2)
        grafo_conexo[t2].add(t1)

    conectadas = set()
    fila = deque([tabela_central])
    while fila:
        atual = fila.popleft()
        if atual not in conectadas:
            conectadas.add(atual)
            fila.extend(grafo_conexo[atual] - conectadas)

    dot = Digraph(comment=f"Snowflake - {nome_base}", format="png")
    dot.attr(rankdir='TB', nodesep='0.3', ranksep='2.0')
    dot.attr('node', shape='box', style='filled,bold', fontname='Arial', fontsize='12')

    dot.node("Legenda", label="""<
        <TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="4" WIDTH="220" BGCOLOR="white">
            <TR><TD><B>Legenda de Relacionamentos</B></TD></TR>
            <TR><TD>→  Filtro da tabela de origem (from) para a tabela de destino (to)</TD></TR>
            <TR><TD>↔️  Filtro bidirecional (ambas direções)</TD></TR>
            <TR><TD><B>1</B>   Um</TD></TR>
            <TR><TD><B>*</B>   Muitos</TD></TR>
        </TABLE>
    >""", shape="plaintext")

    for tabela in conectadas:
        cor = 'gold' if tabela == tabela_central else 'lightyellow'
        dot.node(tabela, fillcolor=cor)

    for _, row in df.iterrows():
        t1, c1 = extrair_tabela_coluna(row['from'])
        t2, c2 = extrair_tabela_coluna(row['to'])
        if t1 in conectadas and t2 in conectadas:
            direcao_simbolo = "↔️" if row["isBidirectional"] else "→"
            cardinalidade = f"{row['fromCardinality']} {direcao_simbolo} {row['toCardinality']}"
            label = f"""<
                <TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="4" WIDTH="100" BGCOLOR="white">
                    <TR>
                        <TD WIDTH="100"><FONT POINT-SIZE="9">{t1}</FONT><BR/><FONT POINT-SIZE="11"><B>{c1}</B></FONT></TD>
                    </TR>
                    <TR>
                        <TD WIDTH="100"><B>{cardinalidade}</B></TD>
                    </TR>
                    <TR>
                        <TD WIDTH="100"><FONT POINT-SIZE="9">{t2}</FONT><BR/><FONT POINT-SIZE="11"><B>{c2}</B></FONT></TD>
                    </TR>
                </TABLE>
            >"""
            style = "bold" if row["isActive"] else "dashed"
            dir_type = "both" if row["isBidirectional"] else "forward"
            dot.edge(t1, t2, label=label, style=style, dir=dir_type)

    os.makedirs(pasta_documentacao, exist_ok=True)
    caminho_imagem = os.path.join(pasta_documentacao, f"{nome_base} - DIAGRAMA.png")
    dot.render(caminho_imagem.replace(".png", ""), cleanup=True)

    return caminho_imagem