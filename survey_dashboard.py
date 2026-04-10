import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go


st.set_page_config(page_title="Dashboard de Clima Organizacional", layout="wide")

# ======================================================
# CARREGAMENTO DOS DADOS
# ======================================================

@st.cache_data
def load_data():
    """Carrega os arquivos Excel de diferentes caminhos possíveis"""
    paths = [
        ('arquivos/', 'Pasta arquivos'),
        ('', 'Raiz do repositório'),
    ]
    
    for path_prefix, desc in paths:
        try:
            surveys = pd.read_excel(f'{path_prefix}surveys.xlsx')
            categories = pd.read_excel(f'{path_prefix}categories.xlsx')
            questions = pd.read_excel(f'{path_prefix}questions.xlsx')
            responses = pd.read_excel(f'{path_prefix}responses.xlsx')
            answers = pd.read_excel(f'{path_prefix}answers.xlsx')
            return surveys, categories, questions, responses, answers
        except (FileNotFoundError, Exception):
            continue
    
    # Se nenhum caminho funcionou
    raise FileNotFoundError("❌ Arquivos Excel não encontrados em nenhum caminho!")

try:
    surveys, categories, questions, responses, answers = load_data()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

# ======================================================
# PROCESSAMENTO
# ======================================================

# Limpa e converte likert_value para numérico
answers['likert_value'] = pd.to_numeric(answers['likert_value'], errors='coerce')

# Filtra apenas respostas Likert (sem NaN)
likert_answers = answers[answers['likert_value'].notna()].copy()

# Filtra respostas abertas
open_answers = answers[answers['open_ended_response'].notna()].copy()

# Merge: answers -> questions -> categories
data = likert_answers.merge(
    questions[['id', 'question_text', 'category_id']], 
    left_on='question_id', 
    right_on='id', 
    how='left',
    suffixes=('', '_q')
)

data = data.merge(
    categories[['id', 'name']], 
    left_on='category_id', 
    right_on='id', 
    how='left',
    suffixes=('', '_cat')
)

# Remove linhas sem categoria
data = data[data['name'].notna()].copy()

# ======================================================
# DASHBOARD
# ======================================================

st.title("📊 Dashboard de Clima Organizacional")
st.markdown("Pesquisa de Clima 2026")
st.markdown("---")

# KPIs Principais
col1, col2, col3 = st.columns(3)

media_geral = data['likert_value'].mean()
col1.metric("📈 Média Geral", f"{media_geral:.2f}/5.0", delta=f"{(media_geral/5)*100:.1f}%")

satisfeitos = (data['likert_value'] >= 4).sum()
taxa_satisfacao = (satisfeitos / len(data) * 100) if len(data) > 0 else 0
col2.metric("😊 Satisfeitos (4-5)", f"{taxa_satisfacao:.1f}%", delta=f"{satisfeitos} votos")

insatisfeitos = (data['likert_value'] <= 2).sum()
taxa_insatisfacao = (insatisfeitos / len(data) * 100) if len(data) > 0 else 0
col3.metric("😞 Insatisfeitos (1-2)", f"{taxa_insatisfacao:.1f}%", delta=f"{insatisfeitos} votos")

st.write(f"**Total de respostas:** {len(data)} | **Total de respondentes:** {data['response_id'].nunique()}")
st.markdown("---")

# ======================================================
# 1. DISTRIBUIÇÃO GERAL
# ======================================================

st.subheader("1️⃣ Como as Pessoas Responderam")

dist = data['likert_value'].value_counts().sort_index()
dist_percent = (dist / len(data) * 100).round(2)

fig_dist = go.Figure(data=[
    go.Bar(
        x=['😞 Ruim (1)', '😐 Fraco (2)', '😑 Neutro (3)', '🙂 Bom (4)', '😍 Ótimo (5)'],
        y=dist_percent.values,
        text=[f"{v:.1f}%" for v in dist_percent.values],
        textposition="outside",
        marker=dict(
            color=['#d73027', '#fc8d59', '#fee090', '#91bfdb', '#4575b4']
        ),
        hovertemplate='<b>%{x}</b><br>%{y:.1f}%<br>%{customdata} votos<extra></extra>',
        customdata=dist.values
    )
])

fig_dist.update_layout(
    title="Percentual de Respostas por Nota",
    xaxis_title="Escala de Satisfação",
    yaxis_title="Percentual (%)",
    height=400,
    showlegend=False,
    template='plotly_dark'
)

st.plotly_chart(fig_dist, use_container_width=True)
st.markdown("---")

# ======================================================
# 2. MÉDIA POR CATEGORIA
# ======================================================

st.subheader("2️⃣ Satisfação por Categoria")

media_cat = data.groupby('name')['likert_value'].agg(['mean', 'count']).reset_index()
media_cat.columns = ['Categoria', 'Média', 'Respostas']
media_cat = media_cat.sort_values('Média', ascending=False)

# Cores baseadas na média
colors = ['#4575b4' if x >= 4 else '#91bfdb' if x >= 3 else '#fee090' if x >= 2 else '#d73027' for x in media_cat['Média']]

fig_cat = go.Figure(data=[
    go.Bar(
        y=media_cat['Categoria'],
        x=media_cat['Média'],
        orientation='h',
        marker=dict(color=colors),
        text=media_cat['Média'].round(2),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Média: %{x:.2f}<br>Respostas: %{customdata}<extra></extra>',
        customdata=media_cat['Respostas']
    )
])

fig_cat.update_layout(
    title='Ranking de Categorias',
    xaxis_title='Média (1-5)',
    yaxis_title='Categoria',
    height=400,
    xaxis_range=[0, 5.1],
    template='plotly_dark'
)

st.plotly_chart(fig_cat, use_container_width=True)

st.markdown("---")

# ======================================================
# 3. TOP 5 MELHORES E PIORES PERGUNTAS
# ======================================================

st.subheader("3️⃣ Ranking de Perguntas")

media_perg = data.groupby('question_text')['likert_value'].agg(['mean', 'count']).reset_index()
media_perg.columns = ['Pergunta', 'Média', 'Respostas']
media_perg = media_perg.sort_values('Média')

col_top, col_worst = st.columns(2)

with col_top:
    st.write("**🔝 Top 5 Melhores Avaliadas**")
    top_5 = media_perg.tail(5).sort_values('Média', ascending=True)
    
    top_5_display = top_5.copy()
    top_5_display['Pergunta'] = top_5_display['Pergunta'].str[:60] + '...'
    
    colors_top = ['#4575b4' if x >= 4 else '#91bfdb' for x in top_5['Média']]
    
    fig_top = go.Figure(data=[
        go.Bar(
            y=top_5_display['Pergunta'],
            x=top_5_display['Média'],
            orientation='h',
            marker=dict(color=colors_top),
            text=top_5['Média'].round(2),
            textposition='outside'
        )
    ])
    fig_top.update_layout(height=300, xaxis_range=[0, 5.1], showlegend=False, template='plotly_dark')
    st.plotly_chart(fig_top, use_container_width=True)

with col_worst:
    st.write("**⚠️ Top 5 Piores Avaliadas**")
    worst_5 = media_perg.head(5)
    
    worst_5_display = worst_5.copy()
    worst_5_display['Pergunta'] = worst_5_display['Pergunta'].str[:60] + '...'
    
    colors_worst = ['#d73027' if x < 2 else '#fc8d59' for x in worst_5['Média']]
    
    fig_worst = go.Figure(data=[
        go.Bar(
            y=worst_5_display['Pergunta'],
            x=worst_5_display['Média'],
            orientation='h',
            marker=dict(color=colors_worst),
            text=worst_5['Média'].round(2),
            textposition='outside'
        )
    ])
    fig_worst.update_layout(height=300, xaxis_range=[0, 5.1], showlegend=False, template='plotly_dark')
    st.plotly_chart(fig_worst, use_container_width=True)

st.markdown("---")

# ======================================================
# 4. INDICADORES DE RISCO
# ======================================================

st.subheader("4️⃣ Áreas Críticas que Precisam de Atenção")

col_risk1, col_risk2 = st.columns(2)

# Categorias críticas
cat_criticas = media_cat[media_cat['Média'] < 3]
with col_risk1:
    st.write("**Categorias com Média < 3.0**")
    if len(cat_criticas) > 0:
        st.warning(f"⚠️ {len(cat_criticas)} categoria(s) crítica(s)")
        for _, row in cat_criticas.iterrows():
            st.write(f"• **{row['Categoria']}**: {row['Média']:.2f}/5")
    else:
        st.success("✅ Nenhuma categoria crítica!")

# Perguntas críticas
perg_criticas = media_perg[media_perg['Média'] < 2.5]
with col_risk2:
    st.write("**Perguntas com Média < 2.5**")
    if len(perg_criticas) > 0:
        st.warning(f"⚠️ {len(perg_criticas)} pergunta(s) crítica(s)")
        for _, row in perg_criticas.iterrows():
            texto = row['Pergunta'][:45] + '...'
            st.write(f"• {texto}: {row['Média']:.2f}")
    else:
        st.success("✅ Nenhuma pergunta crítica!")

st.markdown("---")

# ======================================================
# 5. PERGUNTAS ABERTAS DOS COLABORADORES
# ======================================================

st.subheader("5️⃣ O que os Colaboradores Disseram (Comentários)")

if len(open_answers) > 0:
    # Merge com as perguntas para saber qual pergunta aberta é
    open_data = open_answers.merge(
        questions[['id', 'question_text']], 
        left_on='question_id', 
        right_on='id', 
        how='left'
    )
    
    # Agrupa por pergunta aberta
    pergunta_aberta = open_data['question_text'].unique()
    
    if len(pergunta_aberta) > 0:
        pergunta_selecionada = st.selectbox(
            "Selecione a pergunta aberta:",
            pergunta_aberta
        )
        
        respostas_selecionadas = open_data[open_data['question_text'] == pergunta_selecionada]
        
        st.write(f"**Total de comentários:** {len(respostas_selecionadas)}")
        st.markdown("---")
        
        # Mostra os comentários
        for idx, row in respostas_selecionadas.iterrows():
            if pd.notna(row['open_ended_response']):
                st.write(f"💬 {row['open_ended_response']}")
                st.markdown("---")
    else:
        st.info("Nenhuma pergunta aberta encontrada")
else:
    st.info("Nenhum comentário foi fornecido pelos colaboradores")

st.markdown("---")

# ======================================================
# 6. TABELA DETALHADA DE PERGUNTAS
# ======================================================

st.subheader("6️⃣ Detalhamento por Pergunta")

tabela = media_perg.sort_values('Média', ascending=False).reset_index(drop=True)
tabela['Rank'] = range(1, len(tabela) + 1)
tabela['Média'] = tabela['Média'].round(2)

tabela = tabela[['Rank', 'Pergunta', 'Média', 'Respostas']]

tabela_display = tabela.copy()
tabela_display['Pergunta'] = tabela_display['Pergunta'].str[:80]

st.dataframe(tabela_display, use_container_width=True, hide_index=True)

st.markdown("---")

# ======================================================
# 7. RESUMO EXECUTIVO
# ======================================================

st.subheader("📄 Resumo Executivo para a Liderança")

n_respondentes = data['response_id'].nunique()
taxa_neutro = ((data['likert_value'] == 3).sum() / len(data) * 100)

# Melhor e pior categoria
melhor_cat = media_cat.sort_values('Média', ascending=False).iloc[0]
pior_cat = media_cat.sort_values('Média').iloc[0]

resumo_text = f"""
### 📊 Indicadores Principais

| Métrica | Valor |
|---------|-------|
| **Média Geral** | {media_geral:.2f}/5.0 |
| **Respondentes** | {n_respondentes} pessoas |
| **Total de Respostas** | {len(data)} respostas |
| **Satisfeitos (4-5)** | {taxa_satisfacao:.1f}% |
| **Neutros (3)** | {taxa_neutro:.1f}% |
| **Insatisfeitos (1-2)** | {taxa_insatisfacao:.1f}% |

### 🎯 Destaques

**✅ Melhor categoria:** {melhor_cat['Categoria']} ({melhor_cat['Média']:.2f}/5)

**⚠️ Pior categoria:** {pior_cat['Categoria']} ({pior_cat['Média']:.2f}/5)

**💬 Total de comentários:** {len(open_answers)} respostas abertas

### 💡 Recomendações

1. **Fortalecer:** Replicar boas práticas de {melhor_cat['Categoria']}
2. **Agir:** Investigar causas em {pior_cat['Categoria']}
3. **Escutar:** Analisar os comentários dos colaboradores
4. **Próximas etapas:** Criar plano de ação baseado nesses insights
"""

st.markdown(resumo_text)

st.markdown("---")

# ======================================================
# EXPORTAR
# ======================================================

st.subheader("💾 Exportar Dados")

col_export1, col_export2 = st.columns(2)

with col_export1:
    csv_perg = tabela.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📥 Baixar Análise por Pergunta (CSV)",
        data=csv_perg,
        file_name="analise_perguntas.csv",
        mime="text/csv"
    )

with col_export2:
    csv_cat = media_cat.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📥 Baixar Análise por Categoria (CSV)",
        data=csv_cat,
        file_name="analise_categorias.csv",
        mime="text/csv"
    )

st.markdown("---")
st.caption("Dashboard gerado automaticamente - Atualizar página para recarregar dados")