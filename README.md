# 深度合成信息风险演化与多智能体仿真复现材料

本仓库整理论文第 4.4 节多智能体仿真的代码、输入数据、归档输出和配套图表，用于复现模型运行并比较机制方向与情景差异。

## 内容

- src/abm_final_v2.py：仿真模型和实验设计。
- src/verify_reproduction.py：将重新运行的结果与归档输出逐文件校验。
- data/event_relation_timing_encoding_30.csv：30 条核心独立事件的关系及时序编码。
- data/case_calibration_30.csv：模型案例参数输入。
- data/relation_scene_statistics.csv：关系类型与风险场景的统计结果。
- reference_outputs/：校验脚本所需的 9 个归档 CSV/JSON 输出。
- figures/：模型流程图、核验时差敏感性、智能体规模收敛和情景比较图；模型流程图附 Mermaid 源文件。

## 环境与运行

需要 Python 3.10 或更高版本。

    python -m pip install -r requirements.txt
    python src/abm_final_v2.py --cases data/case_calibration_30.csv --outdir reproduced_output
    python src/verify_reproduction.py --expected reference_outputs --actual reproduced_output

校验脚本对 9 个归档输出进行 SHA-256 比较。所有文件均显示 MATCH，表示本次运行与仓库归档结果逐字节一致。

## 数据与解释边界

30 条核心事件用于关系编码、案例结构比较和时序分析。案例编码为参数关系组合及有序范围提供经验锚定，不等于对仿真行为系数的统计估计。

仿真用于比较机制方向和设定情景差异，不估计现实事件发生概率、现实个体行为系数或政策因果效应。模型输出中的暴露率、错误采信率和现实行动率是模型内部结果，不能解释为现实发生率。统计结果仅适用于本仓库的编码样本及其分析设计。

事件编码表提供公开来源链接供追溯；来源网页及其内容的权利归原发布方所有，使用者应遵守来源条款。

## 引用

引用格式见 CITATION.cff。目前尚无 DOI。

## 许可

本仓库未附许可证。公开可见不代表已授予复制、修改或再分发权限；如需复用，请先联系作者确认。
