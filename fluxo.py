import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
from matplotlib.path import Path

# Configurar o tamanho da figura
plt.figure(figsize=(12, 8))
plt.axis('off')

# Definir cores
cor_mrp = '#D4E6F1'  # Azul claro
cor_crp = '#D5F5E3'  # Verde claro
cor_analise = '#E8DAEF'  # Roxo claro
cor_borda = '#2C3E50'  # Azul escuro
cor_texto = '#34495E'  # Cinza escuro
cor_seta = '#7F8C8D'  # Cinza médio

# Criar retângulos para as fases (x, y, largura, altura)
# MRP - lado esquerdo
inicializacao_mrp = patches.Rectangle((1, 6), 3, 1.5, facecolor=cor_mrp, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                                      zorder=1)
planejamento_mrp = patches.Rectangle((1, 3.5), 3, 1.5, facecolor=cor_mrp, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                                     zorder=1)
execucao_mrp = patches.Rectangle((1, 1), 3, 1.5, facecolor=cor_mrp, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                                 zorder=1)

# CRP - lado direito
inicializacao_crp = patches.Rectangle((8, 6), 3, 1.5, facecolor=cor_crp, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                                      zorder=1)
planejamento_crp = patches.Rectangle((8, 3.5), 3, 1.5, facecolor=cor_crp, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                                     zorder=1)
ajuste_crp = patches.Rectangle((8, 1), 3, 1.5, facecolor=cor_crp, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                               zorder=1)

# Análise - centro inferior
analise = patches.Rectangle((4.5, 0), 3, 1, facecolor=cor_analise, edgecolor=cor_borda, linewidth=1.5, alpha=0.8,
                            zorder=1)

# Adicionar retângulos ao gráfico
ax = plt.gca()
ax.add_patch(inicializacao_mrp)
ax.add_patch(planejamento_mrp)
ax.add_patch(execucao_mrp)
ax.add_patch(inicializacao_crp)
ax.add_patch(planejamento_crp)
ax.add_patch(ajuste_crp)
ax.add_patch(analise)

# Adicionar texto aos retângulos
plt.text(2.5, 6.75, 'Inicialização do MRP', ha='center', va='center', fontsize=12, fontweight='bold', color=cor_texto)
plt.text(2.5, 4.25, 'Planejamento MRP', ha='center', va='center', fontsize=12, fontweight='bold', color=cor_texto)
plt.text(2.5, 1.75, 'Execução e Controle\ndo MRP', ha='center', va='center', fontsize=12, fontweight='bold',
         color=cor_texto)
# Adicionar "Ordens Atualizadas" dentro da caixa Execução e Controle do MRP
plt.text(2.5, 1.25, 'Ordens Atualizadas', ha='center', va='center', fontsize=10, color=cor_texto)

plt.text(9.5, 6.75, 'Inicialização do CRP', ha='center', va='center', fontsize=12, fontweight='bold', color=cor_texto)
plt.text(9.5, 4.25, 'Planejamento CRP', ha='center', va='center', fontsize=12, fontweight='bold', color=cor_texto)
plt.text(9.5, 1.75, 'Ajuste do\nPlanejamento CRP', ha='center', va='center', fontsize=12, fontweight='bold',
         color=cor_texto)
plt.text(9.5, 1.25, 'Feedback do Usuário', ha='center', va='center', fontsize=10, color=cor_texto)

plt.text(6, 0.5, 'Análise', ha='center', va='center', fontsize=12, fontweight='bold', color=cor_texto)

# Definir pontos para as setas
# MRP vertical
seta_mrp_1 = [(2.5, 6), (2.5, 5)]
seta_mrp_2 = [(2.5, 3.5), (2.5, 2.5)]

# CRP vertical
seta_crp_1 = [(9.5, 6), (9.5, 5)]
seta_crp_2 = [(9.5, 3.5), (9.5, 2.5)]

# MRP para CRP (corrigido: agora sai da caixa Execução e Controle do MRP)
seta_mrp_crp = [(4, 1.75), (8, 6.75)]

# Execução para Análise
seta_exec_analise = [(2.5, 1), (4.5, 0.5)]

# Ajuste para Análise
seta_ajuste_analise = [(9.5, 1), (7.5, 0.5)]

# Análise para Inicialização (feedback)
seta_analise_inic = [(6, 0), (6, -0.5), (0.5, -0.5), (0.5, 6.75), (1, 6.75)]


# Função para criar setas
def criar_seta(pontos, largura=0.05, cor=cor_seta):
    verts = [(pontos[0][0], pontos[0][1])]
    codes = [Path.MOVETO]

    for i in range(1, len(pontos)):
        verts.append((pontos[i][0], pontos[i][1]))
        codes.append(Path.LINETO)

    path = Path(verts, codes)
    patch = patches.PathPatch(path, facecolor='none', edgecolor=cor, linewidth=2, zorder=2)
    ax.add_patch(patch)

    # Adicionar ponta da seta
    ultimo_ponto = pontos[-1]
    penultimo_ponto = pontos[-2]

    # Calcular direção
    dx = ultimo_ponto[0] - penultimo_ponto[0]
    dy = ultimo_ponto[1] - penultimo_ponto[1]

    # Normalizar
    comprimento = np.sqrt(dx ** 2 + dy ** 2)
    if comprimento > 0:
        dx = dx / comprimento
        dy = dy / comprimento

    # Calcular pontos da ponta da seta
    ponta_x = ultimo_ponto[0]
    ponta_y = ultimo_ponto[1]

    # Criar triângulo para a ponta da seta
    ponta = patches.Polygon([
        (ponta_x, ponta_y),
        (ponta_x - dx * 0.2 - dy * 0.1, ponta_y - dy * 0.2 + dx * 0.1),
        (ponta_x - dx * 0.2 + dy * 0.1, ponta_y - dy * 0.2 - dx * 0.1)
    ], closed=True, facecolor=cor, edgecolor=cor, zorder=3)

    ax.add_patch(ponta)


# Adicionar todas as setas
criar_seta(seta_mrp_1)
criar_seta(seta_mrp_2)
criar_seta(seta_crp_1)
criar_seta(seta_crp_2)
criar_seta(seta_mrp_crp)
criar_seta(seta_exec_analise)
criar_seta(seta_ajuste_analise)
criar_seta(seta_analise_inic)

# Adicionar rótulos nas setas (evitando sobreposição)
plt.text(2.8, 5.5, 'BOMs, Estoque', fontsize=10, ha='left', va='center', color=cor_texto)
plt.text(2.8, 3.0, 'Demandas, Fluxo de Caixa', fontsize=10, ha='left', va='center', color=cor_texto)
plt.text(6.0, 4.0, 'Planejamento MRP', fontsize=10, ha='center', va='center', color=cor_texto)
plt.text(9.8, 5.5, 'Demandas por Recursos', fontsize=10, ha='left', va='center', color=cor_texto)
plt.text(9.8, 3.0, 'Planilha CRP', fontsize=10, ha='left', va='center', color=cor_texto)
# Adicionar "Insights para Melhorias" próximo ao fluxo de Análise para Inicialização
plt.text(3.0, -0.3, 'Insights para Melhorias', fontsize=10, ha='center', va='center', color=cor_texto)

# Configurar limites do gráfico
plt.xlim(0, 12)
plt.ylim(-1, 8)

# Adicionar título
plt.title('Fluxo de Informações do Sistema Integrado MRP-CRP', fontsize=16, fontweight='bold', pad=20)

# Salvar a imagem
plt.tight_layout()
plt.savefig('fluxo_mrp_crp.png', dpi=300, bbox_inches='tight')
plt.show()
