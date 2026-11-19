# -*- coding: utf-8 -*-
from pathlib import Path
import csv
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import font_manager

OUT = Path('/mnt/data/fourth_chapter_final_evidence')
OUT.mkdir(parents=True, exist_ok=True)

candidates = ['Noto Sans CJK SC','Noto Sans CJK JP','Source Han Sans SC','Microsoft YaHei','SimHei','PingFang SC','Arial Unicode MS']
available = {f.name for f in font_manager.fontManager.ttflist}
font_name = next((x for x in candidates if x in available), None)
if font_name:
    plt.rcParams['font.family'] = font_name
plt.rcParams['axes.unicode_minus'] = False

# Main Figure 2: no embedded title/caption; manuscript supplies caption and note.
scenes = ['公共事件','身份与业务','平台传播','个体权益','组织业务与声誉']
relations = ['来源关系','身份关系','事件关系','传播关系']
counts = np.array([[2,8,8,10],[5,5,5,0],[5,5,0,5],[0,5,5,5],[4,5,5,2]], dtype=float)
ns = np.array([10,5,5,5,5], dtype=float)
pct = counts / ns[:,None] * 100
fig, ax = plt.subplots(figsize=(8.6,5.0))
ax.imshow(pct, aspect='auto')
ax.set_xticks(np.arange(4)); ax.set_xticklabels(relations, fontsize=10.5)
ax.set_yticks(np.arange(5)); ax.set_yticklabels(scenes, fontsize=10.5)
ax.set_xlabel('可验证关系类型', fontsize=10.5)
ax.set_ylabel('风险场景', fontsize=10.5)
for i in range(5):
    for j in range(4):
        ax.text(j,i,f'{int(counts[i,j])}/{int(ns[i])}\n{pct[i,j]:.0f}%',ha='center',va='center',fontsize=9.5)
plt.tight_layout(rect=[0.03,0.03,0.99,0.99])
fig.savefig(OUT/'图2_五类场景可验证关系失稳分布_600dpi.png',dpi=600,bbox_inches='tight')
fig.savefig(OUT/'图2_五类场景可验证关系失稳分布.svg',bbox_inches='tight')
plt.close(fig)

# Main Figure 3: no embedded title/caption; manuscript supplies caption and note.
rows=[]
with (OUT/'abm_delay_intervention_summary.csv').open(encoding='utf-8-sig') as f:
    rows=list(csv.DictReader(f))
fig, ax = plt.subplots(figsize=(8.4,5.0))
for mult in [0.6,1.0,1.4]:
    rr=sorted([r for r in rows if abs(float(r['intervention_multiplier'])-mult)<1e-9], key=lambda x:int(x['delay_shift']))
    x=np.array([int(r['delay_shift']) for r in rr])
    y=np.array([float(r.get('action_rate_mean', r.get('action_mean'))) for r in rr])
    se=np.array([float(r.get('action_rate_se', r.get('action_se'))) for r in rr])
    ax.plot(x,y,marker='o',label=f'干预强度系数={mult:.1f}')
    ax.fill_between(x,y-1.96*se,y+1.96*se,alpha=0.12)
ax.set_xlabel('核验时差增量（离散时间步）',fontsize=10.5)
ax.set_ylabel('现实风险行动率',fontsize=10.5)
ax.legend(frameon=False)
ax.grid(alpha=0.25)
plt.tight_layout(rect=[0.03,0.03,0.99,0.99])
fig.savefig(OUT/'图3_核验时差与干预强度敏感性_600dpi.png',dpi=600,bbox_inches='tight')
fig.savefig(OUT/'图3_核验时差与干预强度敏感性.svg',bbox_inches='tight')
plt.close(fig)

print('clean main figures generated')
