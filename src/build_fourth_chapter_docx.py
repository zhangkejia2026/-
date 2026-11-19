# -*- coding: utf-8 -*-
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

BASE = Path('/mnt/data')
FIGDIR = BASE / 'fourth_chapter_final_evidence'

FONT_CN = 'Noto Serif CJK SC'
FONT_EN = 'Times New Roman'
FONT_MATH = 'Cambria Math'


def set_east_asia(run, font=FONT_CN):
    run.font.name = FONT_EN
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:ascii'), FONT_EN)
    rFonts.set(qn('w:hAnsi'), FONT_EN)


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)


def set_cell_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(cm * 567)))
    tcW.set(qn('w:type'), 'dxa')


def style_run(run, size=10.5, bold=False, italic=False, font=FONT_CN):
    set_east_asia(run, font)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10 if level == 1 else 6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    style_run(r, size=12 if level == 1 else 11, bold=True)
    return p


def add_body(doc, text, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(0)
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)
    r = p.add_run(text)
    style_run(r, 10.5)
    return p


def add_equation(doc, text, number):
    # Use rendered math images so subscripts, superscripts, Greek letters and
    # equation numbering remain stable across Word/LibreOffice renderers.
    path = FIGDIR / f'equation_{number}.png'
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run()
    widths = {1: 9.5, 2: 15.4, 3: 10.0, 4: 15.4, 5: 11.5}
    if path.exists():
        r.add_picture(str(path), width=Cm(widths.get(number, 13.0)))
    else:
        r.text = text + f'    （{number}）'
        style_run(r, 10.5, font=FONT_MATH)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    style_run(r, 9.5)
    return p


def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    style_run(r, 8.5)
    return p


def add_table(doc, headers, rows, widths=None, font_size=8.2):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    table.autofit = False
    hdr_row = table.rows[0]
    hdr = hdr_row.cells
    trPr = hdr_row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), 'true')
    trPr.append(tblHeader)
    cantSplit = OxmlElement('w:cantSplit')
    trPr.append(cantSplit)
    for j, h in enumerate(headers):
        hdr[j].text = str(h)
        set_cell_shading(hdr[j], '4472C4')
        hdr[j].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if widths:
            set_cell_width(hdr[j], widths[j])
        for p in hdr[j].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                style_run(run, font_size, bold=True)
                run.font.color.rgb = RGBColor(255,255,255)
    for row in rows:
        data_row = table.add_row()
        trPr = data_row._tr.get_or_add_trPr()
        cantSplit = OxmlElement('w:cantSplit')
        trPr.append(cantSplit)
        cells = data_row.cells
        for j, val in enumerate(row):
            cells[j].text = str(val)
            cells[j].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths:
                set_cell_width(cells[j], widths[j])
            for p in cells[j].paragraphs:
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 2 else WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    style_run(run, font_size)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def insert_picture(doc, path, width_cm=15.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r.add_picture(str(path), width=Cm(width_cm))
    return p


def create_chapter_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.3); sec.bottom_margin = Cm(2.3)
    sec.left_margin = Cm(2.5); sec.right_margin = Cm(2.5)

    add_heading(doc, '4 基于结构化事件与多智能体仿真的可验证关系失稳分析', 1)
    add_body(doc, '第三章从理论上提出，可验证关系失稳构成深度合成信息风险演化的基础条件，平台传播和公众认知影响不确定性向现实后果转化，风险线索则在演化过程中持续进入证据关联、场景研判和组织响应。为观察上述机制在现实事件中的具体表现，本文将理论概念转换为可编码、可比较和可复核的数据对象，并在结构化事件分析基础上构造多智能体虚拟仿真。经验分析重点回答三个问题：不同场景中哪些可验证关系更易失稳，传播速度与核验速度的落差如何影响风险后果，证据强化、协同响应和可逆干预能否改变风险演化方向。')

    add_heading(doc, '4.1 数据来源与“关系—场景—时序”分析设计', 2)
    add_body(doc, '本文构建由两类资料组成的结构化案例库。第一类为经过去重、事件边界确认和来源升级后的30条核心独立事件，主要用于关系编码、时序比较和风险演化分析；第二类为56条依据政策规范、反诈提示、业务流程和治理规则形成的场景化风险记录，主要用于补充现实事件覆盖不足的业务环节，考察分析框架在不同信息使用环境中的适用性。场景化记录不等同于现实事件，不进入事件频率、结果比较和治理效果统计。')
    add_body(doc, '核心事件纳入须同时满足四项条件：深度合成或人工智能生成属性具有可追溯依据，事件边界能够独立识别，至少具有一项高可信主来源并尽可能配置交叉来源，能够观察到风险后果或治理响应。实验演示、研究数据集、单纯治理公告、深度合成属性无法确认的争议材料，以及同一事件的重复报道和子记录均不进入核心事件统计。清洗后的核心库包括公共事件10条，身份与业务、平台传播、个体权益、组织业务与声誉各5条。')
    add_body(doc, '数据处理采用人机协同方式。智能工具用于事件实体识别、时间和主体字段抽取、来源链接整理、相似记录发现和候选关系识别；事件归并、来源等级、关系判断及结果解释由研究者依据原始材料和统一规则完成。算法检测结果在研究中首先被视为风险线索，只有与来源、身份、事件和传播证据相互印证后，才进入事件事实判断和风险情报形成过程。')

    headers = ['分析维度','编码内容','核心问题','主要证据']
    rows = [
        ['关系维度R','R1来源、R2身份、R3事件、R4传播','什么关系发生失稳','来源凭证、账号身份、授权记录、现实事件、版本与传播记录'],
        ['失稳形态F','缺失、错配、断裂（可多选）','关系以何种方式失去可验证性','元数据不可得、身份/事件错置、传播版本链断裂'],
        ['场景维度S','公共事件、身份与业务、平台传播、个体权益、组织业务与声誉','关系失稳进入何种信息环境','信息用途、风险对象、业务流程与受影响主体'],
        ['时序维度T','t0形成与传播、t1识别核验与治理、t2结果与关系恢复','风险如何随时间演化','首次传播、首次核验、主要处置及结果记录'],
        ['结果维度Y','现实风险行动、风险控制成效、关系恢复','风险最终产生何种后果','财产、权益、决策和秩序后果及治理后结果'],
    ]
    add_table(doc, headers, rows, widths=[2.7,4.2,3.4,5.5], font_size=8.0)
    add_caption(doc, '表2 “关系—场景—时序”结构化编码框架')

    add_body(doc, '关系维度采用多标签编码，同一事件可以同时存在多类关系失稳；缺失、错配和断裂作为辅助形态标签，也不被处理为互斥类别。场景根据主要风险对象和信息用途确定，可验证关系依据事件事实和证据材料独立判断，以避免由场景标签直接推导关系类型。关系—场景关联采用列联表置换检验：在保持场景规模和关系边际分布不变的条件下随机置换20 000次，以经验p值评估样本内部关联，同时报告Cramér V和Benjamini-Hochberg校正结果。鉴于核心事件数量有限，统计检验仅用于结构诊断，不作总体概率推断。')

    add_heading(doc, '4.2 可验证关系失稳的场景差异', 2)
    add_body(doc, '30条核心事件中，身份关系在28条事件中承压，事件关系和传播关系分别出现在23条和22条记录中，来源关系出现在16条记录中。失稳形态方面，错配在全部核心事件中均可观察，缺失和断裂分别出现在16条和22条事件中。错配由此构成深度合成风险的共同结构，而场景差异更多表现为来源缺失、传播断裂以及多类关系叠加方式的不同。')

    headers = ['可验证关系','事件数','Cramér V','置换检验p','FDR校正p']
    rows = [
        ['来源关系','16/30','0.824','<0.001','<0.001'],
        ['身份关系','28/30','0.378','0.541','0.541'],
        ['事件关系','23/30','0.838','<0.001','<0.001'],
        ['传播关系','22/30','0.892','<0.001','<0.001'],
    ]
    add_table(doc, headers, rows, widths=[3.1,2.3,2.8,3.1,3.1], font_size=8.4)
    add_caption(doc, '表3 可验证关系失稳与风险场景的关联强度')
    add_note(doc, '注：p值由20 000次置换获得；FDR为Benjamini-Hochberg校正。检验用于核心事件库内部结构观察，不推断现实总体分布。')

    insert_picture(doc, FIGDIR/'图2_五类场景可验证关系失稳分布_600dpi.png', 15.3)
    add_caption(doc, '图2 五类场景中可验证关系失稳的分布结构（n=30）')
    add_note(doc, '注：基于30条核心独立事件的多标签关系编码；用于样本内部结构观察，不推断现实发生率。')

    add_body(doc, '来源、事件和传播关系与场景之间呈现较强关联，身份关系则接近跨场景共同条件。公共事件的10条记录均涉及传播关系，8条同时涉及身份或事件关系，说明此类风险常表现为主体、事实或语境错配在传播网络中先行扩散。身份与业务场景的5条事件均涉及来源、身份和事件或业务行为对应关系，而传播关系并非主要矛盾，风险即使未进入公共信息流，也可能直接触发资金和授权后果。')
    add_body(doc, '平台传播场景的5条事件均涉及来源、身份和传播关系，个体权益场景则均涉及身份、事件和传播关系。前者主要通过真实素材、公众身份和平台推荐重组可信外观，后者则在身份和事件错配之外叠加复制传播，使人格与权益损害更难逆转。组织业务与声誉场景同时涉及内部业务判断和外部信任，其中身份、事件关系在5条事件中均出现，4条还涉及来源关系。')
    add_body(doc, '上述结果说明，可验证关系失稳不是脱离情境的统一状态。第三章所讨论的关系失稳只有进入具体信息使用环境后，才获得明确的风险含义；场景不仅影响潜在损害，还决定优先核验哪类证据。这一结果直接对应第五章的证据分层与场景化定级：公共事件应优先恢复事件和传播关系，身份与业务应优先确认主体和授权，平台传播需要持续保存来源与版本链，个体权益则应降低受害者取得身份与传播证据的负担。')

    add_heading(doc, '4.3 风险演化的时序特征与核验时差', 2)
    add_body(doc, '为刻画第三章提出的“风险扩展速度—风险情报形成速度”竞争关系，本文将核心事件划分为t0、t1和t2三个阶段。t0记录内容形成、首次传播及关系失稳状态；t1记录风险识别、证据关联和主要治理行动；t2记录现实风险后果、控制成效及关系恢复。进一步依据主要后果出现前后的识别位置，将核验时差划分为四级：0表示事前或实时阻断，1表示事件早期较快介入，2表示明显传播或损失出现后完成主要核验，3表示重大损失或大范围扩散后才形成有效识别。')
    add_body(doc, '30条核心事件中，4条属于事前或即时识别，11条处于较短时差，9条存在明显滞后，6条在主要损失或广泛传播形成后才完成识别。LastPass、WPP和Ferrari等未遂事件表明，声像拟真度较高并不必然导致现实损失；当组织能够调用异常沟通渠道、共享知识或带外确认时，身份关系可以在业务执行前得到核验。与之相对，Arup视频会议诈骗和声音克隆转账事件中，组织在主要资金操作完成后才确认身份和授权关系异常。')
    add_body(doc, '核验时差并不只是组织行动速度的简单度量。它反映风险线索从检测、传播和业务系统进入证据关联，再转化为风险情报所消耗的时间。当来源凭证缺失、身份授权难以确认、跨平台版本不能关联或跨主体证据无法互认时，风险情报形成速度下降。鉴于现实事件难以控制平台放大、认知偏差和场景损害等其他因素，本文不以事件层时差统计直接推断因果关系，而将其转化为多智能体虚拟仿真的核心参数。')

    add_heading(doc, '4.4 多智能体虚拟仿真与机制敏感性分析', 2)
    add_body(doc, '多智能体模型并非另行引入一套外部机制，而是对第三章风险演化过程的计算化映射。内容生成者、平台、用户、核验组织和治理协调者分别对应生成能力、平台传播、公众认知、组织响应和制度协同五个机制环节。模型按照ODD协议描述目的、主体、状态变量、过程、参数和实验设置，以提高透明度和可复现性[27]；用户能力异质性和信息传播状态转换参考已有假信息多主体建模思路[28-29]。')

    headers = ['智能体','主要状态','行为规则','第三章对应机制']
    rows = [
        ['内容生成/攻击者','R1—R4关系失稳、潜在损害D','生成或重构风险对象，初始化关系失稳组合','3.2生成能力与关系失稳'],
        ['平台智能体','可见性A、干预G','按推荐和治理状态分配曝光机会','3.3平台传播'],
        ['用户智能体','易感性S、信息素养L、认知偏差B','接触后决定采信、再传播及现实行动','3.4公众认知'],
        ['核验组织','关系证据Er、核验时差τ','关联多源证据，更新未解决关系失稳和风险判断','3.5风险情报形成'],
        ['治理协调者','协同C、证据阈值、损害阈值','实施观察、可逆临时干预、证据驱动升级或退出','3.6制度协同与反馈'],
    ]
    add_table(doc, headers, rows, widths=[3.0,4.2,5.4,3.3], font_size=7.8)
    add_caption(doc, '表4 多智能体模型的主体、状态与机制映射')

    add_body(doc, '设Rr表示第r类关系的初始失稳状态，Er(t)表示该关系在时间t的证据充分程度。未解决关系失稳定义为：')
    add_equation(doc, 'U(t)=Σ₍r=1…4₎ R_r[1−E_r(t)] / Σ₍r=1…4₎R_r', 1)
    add_body(doc, 'Er(t)在关系特定的核验时差τr之前为0，之后采用指数饱和函数描述证据逐步形成：')
    add_equation(doc, 'E_r(t)=0，t<τ_r；  E_r(t)=1−exp{−q_r(1+αC)(t−τ_r+1)/κ}，t≥τ_r', 2)
    add_body(doc, '平台只决定信息可见性，避免在用户再传播概率中重复计量。时间t由活跃传播者形成的新接触数服从：')
    add_equation(doc, 'K_t ~ Poisson[λ₀·A_t·(1−G_t)·I_t]', 3)
    add_body(doc, '用户采信、再传播和现实行动概率均采用Logistic函数。采信与再传播由个体易感性、信息素养、认知偏差和未解决关系失稳共同决定；现实行动在已采信用户中进一步受到场景潜在损害D影响。平台变量A不再进入用户再传播方程。')
    add_equation(doc, 'P_bel=σ(β₀+β_sS_i+β_bB+β_uU−β_lL_i)； P_share=σ(γ₀+γ_sS_i+γ_bB+γ_uU−γ_lL_i)', 4)
    add_equation(doc, 'P_act=σ(δ₀+δ_sS_i+δ_bB+δ_dD+δ_uU−δ_lL_i)', 5)
    add_body(doc, '模型以现实风险行动率Y=N⁻¹ΣAction_i作为下游结果，不再使用缺乏经验依据的加权风险指数；关系恢复率由RR(T)=1−U(T)直接计算。治理协调者依据证据可信度与潜在损害实施分级行动：证据不足但潜在损害较高时，可启动最长3个时间步的可逆临时干预；证据达到阈值后再升级为实质干预，未达到阈值则自动退出。')

    p_conv = add_body(doc, '智能体数量通过多规模Monte Carlo收敛检验确定。本文依次设置N=1 000、2 000、4 000、8 000和16 000，初始风险种子固定为用户总数的0.5%，每个规模重复500次，并以N=16 000作为参照。N=4 000时，暴露率、错误采信率、现实行动率和关系恢复率四项均值相对参照的最大绝对差为0.0084，低于0.01的预设阈值，因此选取N=4 000作为满足稳定性要求的最小规模。模型参数用于表达理论机制的方向和相对强度，不是由30条事件估计得到的现实行为系数；完整参数、原始输出和代码作为复现材料提供。')

    headers = ['情景','暴露率','错误采信率','现实行动率','关系恢复率','相对控制成效']
    rows = [
        ['基准','30.08%','18.02%','7.80%','95.79%','72.47%'],
        ['平台放大','63.85%','38.93%','17.09%','95.79%','45.88%'],
        ['认知偏差增强','47.46%','30.08%','13.64%','95.79%','59.02%'],
        ['响应滞后','56.79%','35.54%','15.76%','91.14%','44.38%'],
        ['关系证据强化','13.11%','7.56%','3.22%','99.40%','88.65%'],
        ['协同与可逆干预','14.54%','8.71%','3.75%','97.41%','86.78%'],
    ]
    add_table(doc, headers, rows, widths=[3.2,2.3,2.5,2.5,2.5,2.7], font_size=8.0)
    add_caption(doc, '表5 多智能体虚拟仿真的机制情景结果')
    add_note(doc, '注：N=4 000，每类情景500次Monte Carlo重复；表中为均值。相对控制成效=1−治理情景行动率/配对无治理情景行动率。结果用于机制方向比较，不代表现实发生概率或政策因果效应。')

    add_body(doc, '情景结果显示，平台放大和响应滞后对风险演化具有明显的方向性影响。平台放大情景的现实行动率由基准情景的7.80%升至17.09%，相对控制成效下降至45.88%；核验时差增加3个时间步后，现实行动率升至15.76%，关系恢复率降至91.14%。认知偏差增强同样提高错误采信和现实行动，但其影响低于平台放大与响应滞后。平台放大和认知偏差情景未改变证据累积参数，因此终期关系恢复率与基准基本一致；这说明“最终能够澄清”并不意味着前期风险后果可以同步撤回。')
    add_body(doc, '关系证据强化通过提高证据质量并缩短核验时差，使现实行动率降至3.22%，关系恢复率提高至99.40%；协同与可逆干预情景将早期传播摩擦、跨主体证据整合和后续证据驱动治理结合，现实行动率降至3.75%，相对控制成效达到86.78%。这两类情景表明，治理成效不仅取决于最终是否删除内容，更取决于风险线索能否较快转化为可信证据并进入行动。')

    insert_picture(doc, FIGDIR/'图3_核验时差与干预强度敏感性_600dpi.png', 15.3)
    add_caption(doc, '图3 核验时差与干预强度对现实风险后果的敏感性')
    add_note(doc, '注：N=4 000，每组400次Monte Carlo重复；曲线为均值，阴影为95% Monte Carlo置信区间。仿真用于机制敏感性分析，不作现实因果效应估计。')
    add_body(doc, '图3进一步显示，核验时差从0增加至6个时间步时，三种干预强度下的现实行动率均持续上升；提高干预强度可以降低风险，但无法完全抵消较长时差形成的累积效应。对平台接触率、证据质量、临时干预强度和现实行动阈值实施±20%的单因素扰动后，平台放大增加风险、证据强化和干预增强降低风险的方向保持稳定。由此，模型支持将治理窗口置于末端治理强度之前：前端证据组织和风险情报形成速度，可能比事后提高处置强度更具机制意义。')

    add_heading(doc, '4.5 典型事件对照与分析框架深化', 2)
    add_body(doc, '结构化事件与虚拟仿真共同表明，相似的深度合成技术可以产生完全不同的现实结果，关键差异不只在内容拟真度，而在于关系证据能否被及时调用以及风险情报形成速度能否赶上风险扩展。身份与业务场景中，Arup视频会议诈骗与Ferrari、WPP、LastPass未遂事件构成较清晰的对照。前者利用高度拟真的视频会议同时伪装多名组织成员，风险直接进入资金流程，组织在主要损失形成后才确认身份和授权关系异常；后三类事件同样采用高管身份仿冒和声音克隆，但员工通过异常沟通渠道、共享知识或二次确认识别问题，使风险停留在未遂阶段。')
    add_body(doc, '公共传播与个体权益场景也呈现不同的可逆性。五角大楼“爆炸”AI图像短时获得较高可见性并造成市场波动，但官方快速核验恢复了事件关系，现实后果持续时间相对有限；Taylor Swift非自愿深度伪造图像虽遭平台后续处置，但内容已经大规模复制传播，身份与事件错配造成的人格和权益损害难以通过删除完全逆转。风险治理因而不仅需要判断内容是否真实，还需判断不同场景中的损害是否可逆以及关系证据能否在关键窗口内恢复。')
    add_body(doc, '第四章由此对第三章形成三方面经验深化：可验证关系失稳只有进入具体场景后才获得明确风险含义；传播速度与核验速度的落差是风险由信息异常转化为现实后果的重要条件；治理效果取决于证据质量、风险判断和行动之间能否连续衔接。结构化事件揭示现实机制的具体形态，多智能体仿真则在统一规则下观察平台放大、认知偏差、核验时差和治理程序变化时的方向性结果。二者均不承担现实因果效应的最终证明，但共同支持将深度合成情报治理的重点由末端内容辨伪前移至证据组织、场景研判和风险情报形成过程。')
    add_body(doc, '据此，第五章围绕第四章识别出的三类断点展开：针对关系证据不足建立显式标识、来源凭证和事件证据的分层体系；针对场景损害与证据状态错配实施动态定级和可逆、限时干预；针对核验时差和跨主体信息断裂建立最小必要共享与协同响应，并通过复盘学习持续恢复来源、身份、事件和传播关系的可验证性。')
    return doc


def replace_paragraph_text(paragraph, text, size=10.5, bold=False, center=False, indent=True):
    for run in paragraph.runs:
        run.text = ''
    if not paragraph.runs:
        run = paragraph.add_run()
    else:
        run = paragraph.runs[0]
    run.text = text
    style_run(run, size, bold=bold)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.line_spacing = 1.25
    paragraph.paragraph_format.first_line_indent = Pt(21) if indent else None


def remove_between(doc, start_text, end_text):
    start = None; end = None
    for p in doc.paragraphs:
        t = p.text.strip()
        if start is None and t.startswith(start_text): start = p._p
        elif start is not None and t.startswith(end_text): end = p._p; break
    if start is None or end is None:
        raise RuntimeError('chapter boundaries not found')
    body = doc._body._element
    children = list(body)
    si = children.index(start); ei = children.index(end)
    for child in children[si:ei]:
        body.remove(child)
    return end


def insert_chapter_before(doc, chapter_doc, anchor):
    # Copy text and tables only. Drawing relationship IDs are document-specific;
    # pictures/equations are reinserted in the destination document afterwards.
    for child in list(chapter_doc._body._element):
        if child.tag == qn('w:sectPr'):
            continue
        if child.xpath('.//w:drawing'):
            continue
        anchor.addprevious(deepcopy(child))


def insert_picture_before_anchor(doc, anchor_paragraph, path, width_cm):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r.add_picture(str(path), width=Cm(width_cm))
    anchor_paragraph._p.addprevious(p._p)
    return p


def insert_chapter4_images(doc):
    # Figure 2
    cap2 = next(p for p in doc.paragraphs if p.text.strip().startswith('图2 五类场景中可验证关系失稳'))
    insert_picture_before_anchor(doc, cap2, FIGDIR/'图2_五类场景可验证关系失稳分布_600dpi.png', 15.3)

    # Equations 1-3 are inserted before the explanatory paragraph that follows them.
    anchors = {
        1: next(p for p in doc.paragraphs if p.text.strip().startswith('Er(t)在关系特定的核验时差')),
        2: next(p for p in doc.paragraphs if p.text.strip().startswith('平台只决定信息可见性')),
        3: next(p for p in doc.paragraphs if p.text.strip().startswith('用户采信、再传播和现实行动概率')),
    }
    widths = {1:9.5, 2:15.4, 3:10.0, 4:15.4, 5:11.5}
    for n in (1,2,3):
        insert_picture_before_anchor(doc, anchors[n], FIGDIR/f'equation_{n}.png', widths[n])

    # Equations 4 and 5 occur consecutively before the outcome-definition paragraph.
    outcome = next(p for p in doc.paragraphs if p.text.strip().startswith('模型以现实风险行动率'))
    insert_picture_before_anchor(doc, outcome, FIGDIR/'equation_4.png', widths[4])
    insert_picture_before_anchor(doc, outcome, FIGDIR/'equation_5.png', widths[5])

    # Figure 3
    cap3 = next(p for p in doc.paragraphs if p.text.strip().startswith('图3 核验时差与干预强度'))
    insert_picture_before_anchor(doc, cap3, FIGDIR/'图3_核验时差与干预强度敏感性_600dpi.png', 15.3)


def replace_governance_image(doc):
    p5 = None; p6 = None
    for p in doc.paragraphs:
        if p.text.strip().startswith('5.5 '): p5 = p
        if p.text.strip().startswith('6 '): p6 = p; break
    if p5 is None or p6 is None: return
    body = doc._body._element
    children = list(body)
    start_i = children.index(p5._p); end_i = children.index(p6._p)
    image_p = None
    for child in children[start_i:end_i]:
        if child.tag == qn('w:p') and child.xpath('.//w:drawing'):
            image_p = child
            break
    if image_p is None: return
    body.remove(image_p)
    # Remove the old figure caption(s) left in chapter 5.
    for child in list(body)[start_i:end_i]:
        if child.tag == qn('w:p'):
            txt = ''.join(child.itertext()).strip()
            if txt.startswith('图3') or txt.startswith('图 3') or txt.startswith('图4') or txt.startswith('图 4'):
                body.remove(child)
    # Add the new image directly to the destination document so its relationship ID is valid.
    img_p = doc.add_paragraph()
    img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_run = img_p.add_run()
    img_run.add_picture(str(BASE/'图4_治理机制_无内嵌图号.png'), width=Cm(15.4))
    p6._p.addprevious(img_p._p)

    cap_p = doc.add_paragraph()
    cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_p.paragraph_format.space_before = Pt(3)
    cap_p.paragraph_format.space_after = Pt(5)
    cap_r = cap_p.add_run('图4 面向可验证关系重建的深度合成信息风险情报治理机制')
    style_run(cap_r, 9.5)
    p6._p.addprevious(cap_p._p)


def append_references(doc):
    refs = [
        '[27] Grimm V, Railsback S F, Vincenot C E, et al. The ODD Protocol for Describing Agent-Based and Other Simulation Models: A Second Update to Improve Clarity, Replication, and Structural Realism[J]. Journal of Artificial Societies and Social Simulation, 2020, 23(2): 7.',
        '[28] Franceschi J, Pareschi L, Zanella M. From Agent-Based Models to the Macroscopic Description of Fake-News Spread: The Role of Competence in Data-Driven Applications[J]. Partial Differential Equations and Applications, 2022, 3: 68.',
        '[29] Alassad M, Hussain M N, Agarwal N. Developing an Agent-Based Model to Minimize Spreading of Malicious Information in Dynamic Social Networks[J]. Computational and Mathematical Organization Theory, 2023, 29(3): 487-502.',
    ]

    # Remove stale copies from prior builds.
    for p in list(doc.paragraphs):
        if p.text.strip().startswith(('[27]', '[28]', '[29]')):
            p._element.getparent().remove(p._element)

    p26 = next(p for p in doc.paragraphs if p.text.strip().startswith('[26]'))
    anchor = p26._p
    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(ref)
        style_run(r, 10.5)
        anchor.addnext(p._p)
        anchor = p._p

def build_full_paper(chapter_doc):
    src = BASE/'深度合成信息风险的生成机制与情报治理路径V2-6.docx'
    doc = Document(src)
    anchor = remove_between(doc, '4 ', '5 ')
    insert_chapter_before(doc, chapter_doc, anchor)
    insert_chapter4_images(doc)

    # Update abstract and introduction method/contribution paragraphs.
    abstract_text = ('摘要：[目的/意义] 深度合成技术使数字内容与来源主体、身份对象、现实事件及传播记录之间的对应关系更易发生缺失、错配和断裂，真实性判断由内容真伪识别进一步转向多源关系核验。本文引入“可验证关系失稳”解释信息风险演化，并讨论风险情报形成机制。'
                     '[方法/过程] 构建“可验证关系失稳—风险演化与风险情报形成并行推进—治理响应—可验证关系重建”的分析框架，建立30条核心独立事件与56条场景化风险记录组成的结构化案例库，采用“关系—场景—时序”编码、置换关联分析和多智能体虚拟仿真考察风险机制。'
                     '[结果/结论] 事件分析显示，来源、事件和传播关系的失稳具有明显场景差异，身份关系则呈跨场景共通性；仿真表明，平台放大、认知偏差和核验时差会提高错误采信与现实风险行动，关系证据强化、协同响应和可逆干预能够缩短治理窗口并提高关系恢复。治理应由内容辨伪转向证据分层、动态定级、可逆干预、最小必要共享和复盘学习。')
    replace_paragraph_text(doc.paragraphs[3], abstract_text, 10.5, indent=False)

    p_method = next(p for p in doc.paragraphs if p.text.startswith('基于此，本文从可验证关系失稳出发'))
    replace_paragraph_text(p_method, '基于此，本文从可验证关系失稳出发考察深度合成信息风险的演化机制，并将风险情报作为连接风险识别与治理行动的重要环节。在理论分析基础上，构建由30条核心独立事件和56条场景化风险记录组成的案例库，采用“关系—场景—时序”结构化编码、置换关联分析和多智能体虚拟仿真，对关系失稳的场景差异、核验时差及治理机制开展经验观察与敏感性分析，并进一步提出面向可验证关系重建的情报治理机制。', 10.5)
    p_contrib = next(p for p in doc.paragraphs if p.text.startswith('本文的推进主要体现在'))
    replace_paragraph_text(p_contrib, '本文的推进主要体现在三个方面：一是将分析单位由单一内容真实性推进到数字内容、现实对象与多源证据之间的可验证关系，解释“局部真实、整体失真”条件下风险如何形成；二是将风险线索向风险情报的转化纳入分析，讨论多源核验、场景研判与组织响应如何把分散信息转化为具有等级、时效和行动指向的组织判断；三是形成计算辅助的“关系—场景—时序”结构化事件分析，并以多智能体虚拟仿真观察平台放大、认知偏差、核验时差和治理程序变化时的机制敏感性。案例与仿真均不承担现实总体因果效应推断。', 10.5)

    # Update conclusion paragraphs for coherence.
    p136 = next(p for p in doc.paragraphs if p.text.startswith('结构化案例材料为上述框架提供'))
    replace_paragraph_text(p136, '结构化事件与多智能体仿真为上述框架提供了相互补充的经验支持。30条核心事件表明，来源、事件和传播关系的失稳具有明显场景差异，身份关系则呈现跨场景共通性；风险传播与事实核验之间的时间落差，决定了关系失稳是否在关键窗口内进入现实行动。虚拟仿真进一步显示，在统一规则下提高平台放大、认知偏差或核验时差会增加错误采信与现实风险行动，而关系证据强化、协同响应和可逆干预能够降低风险并提高关系恢复。上述结果用于机制方向与敏感性观察，不构成现实政策因果效应的证明。', 10.5)
    p138 = next(p for p in doc.paragraphs if p.text.startswith('本文仍存在明确的证据边界'))
    replace_paragraph_text(p138, '本文仍存在明确的证据边界。核心事件库采用目的性纳入而非概率抽样，场景化风险记录只用于框架覆盖与情报任务比较；多智能体模型的行为系数和参数区间用于表达理论机制的方向与相对强度，不是由案例资料估计得到的现实行为概率。后续研究可进一步引入平台级传播日志、组织业务数据、独立编码和纵向事件材料，对关系证据累积、风险情报形成速度及治理程序的实际效果进行校准和检验。总体而言，生成式人工智能环境下的真实性问题正在由“内容是真是假”进一步转向“关系能否被验证”，这也构成深度合成信息风险研究由内容辨伪走向计算情报治理的重要方向。', 10.5)

    replace_governance_image(doc)
    append_references(doc)
    return doc


if __name__ == '__main__':
    chapter = create_chapter_doc()
    standalone = BASE/'第四章_可验证关系失稳与多智能体仿真_最优重构版.docx'
    chapter.save(standalone)
    full = build_full_paper(chapter)
    full_path = BASE/'深度合成信息风险的生成机制与情报治理路径V2-7_第四章最优重构版.docx'
    full.save(full_path)
    print(standalone)
    print(full_path)
