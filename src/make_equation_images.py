# -*- coding: utf-8 -*-
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib import font_manager

OUT = Path('/mnt/data/fourth_chapter_final_evidence')
OUT.mkdir(exist_ok=True)

# Use STIX for mathematical glyph coverage.
plt.rcParams['mathtext.fontset'] = 'stix'
plt.rcParams['font.family'] = 'STIXGeneral'

formulas = {
1: r'$U(t)=\dfrac{\sum_{r=1}^{4}R_r\,[1-E_r(t)]}{\sum_{r=1}^{4}R_r}$',
2: r'$E_r(t)=0\quad(t<\tau_r);\qquad E_r(t)=1-\exp\!\left[-\dfrac{q_r(1+\alpha C)(t-\tau_r+1)}{\kappa}\right]\quad(t\geq\tau_r)$',
3: r'$K_t\sim\mathrm{Poisson}\!\left[\lambda_0 A_t(1-G_t)I_t\right]$',
4: r'$P_i^{\mathrm{bel}}(t)=\sigma(\beta_0+\beta_sS_i+\beta_bB+\beta_uU(t)-\beta_lL_i),\quad P_i^{\mathrm{share}}(t)=\sigma(\gamma_0+\gamma_sS_i+\gamma_bB+\gamma_uU(t)-\gamma_lL_i)$',
5: r'$P_i^{\mathrm{act}}(t)=\sigma(\delta_0+\delta_sS_i+\delta_bB+\delta_dD+\delta_uU(t)-\delta_lL_i)$',
}

sizes = {1:(8.4,0.62),2:(11.0,0.68),3:(8.3,0.62),4:(12.0,0.68),5:(9.5,0.62)}
for n, formula in formulas.items():
    w,h=sizes[n]
    fig=plt.figure(figsize=(w,h), facecolor='none')
    ax=fig.add_axes([0,0,1,1]); ax.axis('off')
    ax.text(0.50,0.50,formula,ha='center',va='center',fontsize=17,color='black')
    ax.text(0.985,0.50,f'({n})',ha='right',va='center',fontsize=13,color='black')
    fig.savefig(OUT/f'equation_{n}.png',dpi=300,transparent=True,bbox_inches='tight',pad_inches=0.03)
    fig.savefig(OUT/f'equation_{n}.svg',transparent=True,bbox_inches='tight',pad_inches=0.03)
    plt.close(fig)
print('equation images generated')
