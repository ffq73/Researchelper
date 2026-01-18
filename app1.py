import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import docx
import openpyxl
import re
import io
import dashscope
import json
import time

# ==========================================
# 0. 全局配置
# ==========================================
st.set_page_config(page_title="行研 Copilot Ultimate", layout="wide", page_icon="🚀")

# 设置中文字体 (兼容 Windows/Linux)
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial', 'Microsoft YaHei', 'PingFang SC']
plt.rcParams['axes.unicode_minus'] = False

# 初始化 Session State
if 'ai_config' not in st.session_state: st.session_state['ai_config'] = None
if 'df_cache' not in st.session_state: st.session_state['df_cache'] = None
if 'compliance_results' not in st.session_state: st.session_state['compliance_results'] = []

# ==========================================
# 1. 核心解析器 (高容错版)
# ==========================================

def clean_text(text):
    """清洗：去重空白符"""
    if not text: return ""
    return "".join(str(text).split())

def split_segments(full_text):
    """分词：按标点切分，保留长句"""
    segments = re.split(r'[。；！？\n]+', str(full_text))
    return set([clean_text(s) for s in segments if len(clean_text(s)) > 2])

@st.cache_data(show_spinner=False)
def get_docx_text(file):
    """解析 Word (含复杂表格修复)"""
    try:
        doc = docx.Document(file)
    except: return set(), ""
    
    txt = []
    # 段落
    for p in doc.paragraphs: txt.append(p.text)
    # 表格 (暴力容错读取)
    for t in doc.tables:
        for r in t.rows:
            try:
                for c in r.cells: txt.append(c.text)
            except: 
                # XML 暴力读取兜底
                try:
                    for cell in r._element.tc_lst:
                        for p in cell.p_lst:
                            nodes = p.xpath('.//w:t')
                            txt.append("".join([n.text for n in nodes if n.text]))
                except: pass
    raw = "\n".join(txt)
    return split_segments(raw), raw

@st.cache_data(show_spinner=False)
def get_pptx_text(file):
    """解析 PPT"""
    try:
        prs = Presentation(file)
    except: return set(), ""
    txt = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"): txt.append(shape.text)
            if shape.has_table:
                for r in shape.table.rows:
                    for c in r.cells: txt.append(c.text)
    raw = "\n".join(txt)
    return split_segments(raw), raw

@st.cache_data(show_spinner=False)
def get_excel_text(file):
    """解析 Excel"""
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
    except: return set(), ""
    txt = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows(values_only=True):
            for c in row:
                if c: txt.append(str(c))
    raw = "\n".join(txt)
    return split_segments(raw), raw

def dispatch_extractor(file):
    if file.name.endswith('.docx'): return get_docx_text(file)
    elif file.name.endswith('.pptx'): return get_pptx_text(file)
    elif file.name.endswith('.xlsx'): return get_excel_text(file)
    return set(), ""

# ==========================================
# 2. 模块：全格式核对 (分批次全覆盖版)
# ==========================================

def run_ai_batch_check(api_key, context, targets):
    """
    AI 批次处理函数：由于 LLM 输出长度有限，我们需要把差异项分批发送
    """
    dashscope.api_key = api_key
    # 构造 Prompt：要求 AI 返回 JSON 格式
    prompt = f"""
    你是一个极其严苛的金融合规审核员。
    【基准事实】(截取部分):
    {context}
    
    【待审核列表】:
    {json.dumps(targets, ensure_ascii=False)}

    【指令】
    请逐条判断【待审核列表】中的内容在【基准事实】中是否有依据。
    不要遗漏任何一条。
    请严格返回 JSON 数组格式 (不要 Markdown)：
    [
        {{"text": "原句", "result": "✅通过/❌存疑", "reason": "简短理由"}}
    ]
    若语义一致或数据匹配，标记通过；若无中生有或数据错误，标记存疑。
    """
    try:
        # 使用 turbo 模型速度较快，如果需要更高精度可换 qwen-plus
        resp = dashscope.Generation.call(model='qwen-turbo', prompt=prompt)
        content = resp.output.text.replace("```json", "").replace("```", "").strip()
        return json.loads(content)
    except Exception as e:
        return [{"text": t, "result": "⚠️API错误", "reason": str(e)} for t in targets]

def module_compliance(api_key):
    st.header("🕵️ 全格式文档核对 (全量覆盖版)")
    st.markdown("支持 Word/Excel/PPT 互查。**策略：** 优先展示完整差异，AI 分批次扫描所有条目，确保 0 遗漏。")
    
    c1, c2 = st.columns(2)
    f1 = c1.file_uploader("1. 基准文件 (Source)", type=['docx','xlsx','pptx'], key="cf1")
    f2 = c2.file_uploader("2. 待测文件 (Target)", type=['docx','xlsx','pptx'], key="cf2")
    
    if f1 and f2:
        with st.spinner("正在解析文档结构..."):
            s1, raw1 = dispatch_extractor(f1)
            s2, raw2 = dispatch_extractor(f2)
            ghosts = list(s2 - s1)
        
        # --- 优化点：AI 介入前，先展示完整差异 ---
        if not ghosts:
            st.success("✅ 完美匹配！无任何差异内容。")
        else:
            st.warning(f"⚠️ 共发现 {len(ghosts)} 处原始内容差异")
            
            # 展示原生差异列表
            with st.expander("📄 查看完整差异清单 (点击展开)", expanded=True):
                st.dataframe(pd.DataFrame(ghosts, columns=["待审核内容"]), use_container_width=True)

            # --- AI 全量扫描 ---
            if st.button("🧠 AI 全量深度判别 (覆盖所有条目)", type="primary", key="btn_ai_full"):
                if not api_key:
                    st.error("请先输入 API Key")
                    return

                # 初始化进度条
                progress_bar = st.progress(0)
                status_text = st.empty()
                all_results = []
                
                # 分批策略：每次处理 20 条，防止 Token 溢出
                BATCH_SIZE = 20
                total_items = len(ghosts)
                
                # 截取基准上下文 (保留前 2.5万字，通常涵盖核心内容)
                # 这里的策略是：保证上下文足够长，而不是为了速度去过度裁剪
                safe_context = raw1[:25000]
                
                for i in range(0, total_items, BATCH_SIZE):
                    batch_targets = ghosts[i : i + BATCH_SIZE]
                    status_text.text(f"AI 正在审核第 {i+1} ~ {min(i+BATCH_SIZE, total_items)} 条，共 {total_items} 条...")
                    
                    # 调用 AI
                    batch_res = run_ai_batch_check(api_key, safe_context, batch_targets)
                    all_results.extend(batch_res)
                    
                    # 更新进度
                    progress_bar.progress(min((i + BATCH_SIZE) / total_items, 1.0))
                
                status_text.text("✅ 审核完成！")
                st.session_state['compliance_results'] = all_results

            # --- 展示 AI 结果 ---
            if st.session_state['compliance_results']:
                st.divider()
                st.subheader("📋 AI 审核报告")
                
                res_df = pd.DataFrame(st.session_state['compliance_results'])
                
                # 高亮显示存疑项
                def highlight_row(row):
                    if "❌" in str(row['result']) or "⚠️" in str(row['result']):
                        return ['background-color: #ffcccc'] * len(row)
                    return [''] * len(row)

                st.dataframe(
                    res_df.style.apply(highlight_row, axis=1), 
                    use_container_width=True,
                    column_config={
                        "text": "原句",
                        "result": "判定结果",
                        "reason": "AI 理由"
                    }
                )
                
                # 统计概览
                fail_count = len([x for x in st.session_state['compliance_results'] if "❌" in x['result']])
                if fail_count > 0:
                    st.error(f"发现 {fail_count} 个高风险存疑项，请重点核对！")
                else:
                    st.success("所有差异项均已通过 AI 语义合规性检查。")

# ==========================================
# 3. 模块：智能制图 (范例仿制版)
# ==========================================

def ai_analyze_chart(api_key, df):
    """AI 图表配置分析"""
    dashscope.api_key = api_key
    data_sample = df.head(3).to_json(orient='records', force_ascii=False)
    prompt = f"""
    分析数据样例: {data_sample}
    给出 Matplotlib 绘图建议。严格返回 JSON:
    {{
        "chart_type": "dual_axis" 或 "bar" 或 "line",
        "x_col": "时间或类别列名",
        "y_primary": ["主轴列名"],
        "y_secondary": ["副轴列名"],
        "title": "建议标题"
    }}
    """
    try:
        resp = dashscope.Generation.call(model='qwen-turbo', prompt=prompt)
        return json.loads(resp.output.text.replace("```json","").replace("```","").strip())
    except: return None

def module_smart_chart_ref(api_key):
    st.header("📊 智能制图 (范例仿制)")
    st.markdown("工作流：上传参考图 -> AI分析数据 -> 调整样式以匹配参考图 -> 导出。")
    
    c1, c2 = st.columns(2)
    ref_img = c1.file_uploader("1. 参考范例 (截图)", type=['png','jpg'], key="ci_1")
    data_file = c2.file_uploader("2. 数据 Excel", type=['xlsx'], key="ci_2")
    
    if ref_img: c1.image(ref_img, caption="目标样式", use_column_width=True)
    
    if data_file and api_key:
        df = pd.read_excel(data_file)
        st.session_state['df_cache'] = df
        
        if st.button("🤖 AI 分析数据结构", key="btn_chart_ai"):
            with st.spinner("AI 正在解析数据维度..."):
                cfg = ai_analyze_chart(api_key, df)
                if cfg:
                    st.session_state['ai_config'] = cfg
                    st.success("分析完成，请下方调整。")
                else: st.error("AI 分析失败")

    if st.session_state['ai_config']:
        cfg = st.session_state['ai_config']
        df = st.session_state['df_cache']
        cols = df.columns.tolist()
        
        st.divider()
        st.subheader("🎨 样式对齐")
        cc1, cc2 = st.columns([1, 2])
        
        with cc1:
            # 交互微调区
            c_type = st.selectbox("图表类型", ["dual_axis", "bar", "line"], index=["dual_axis", "bar", "line"].index(cfg.get('chart_type','bar')), key="s_type")
            c_x = st.selectbox("X轴", cols, index=cols.index(cfg.get('x_col')) if cfg.get('x_col') in cols else 0, key="s_x")
            
            def_y1 = [c for c in cfg.get('y_primary',[]) if c in cols]
            c_y1 = st.multiselect("左轴数据", cols, default=def_y1 if def_y1 else [cols[1]], key="s_y1")
            
            def_y2 = [c for c in cfg.get('y_secondary',[]) if c in cols]
            c_y2 = st.multiselect("右轴数据", cols, default=def_y2, key="s_y2")
            
            st.markdown("---")
            col1 = st.color_picker("主色 (吸取参考图)", "#C00000", key="cp1")
            col2 = st.color_picker("副色", "#FFC000", key="cp2")
            f_size = st.slider("字号", 8, 20, 10, key="fs")
            c_title = st.text_input("标题", value=cfg.get('title','Chart'), key="st")

        with cc2:
            # 绘图逻辑
            plt.rcParams.update({'font.size': f_size})
            fig, ax1 = plt.subplots(figsize=(8, 4.5))
            
            if c_type == "dual_axis":
                ax1.bar(df[c_x], df[c_y1[0]], color=col1, alpha=0.8, label=c_y1[0])
                ax1.set_ylabel(c_y1[0], color=col1, fontweight='bold')
                if c_y2:
                    ax2 = ax1.twinx()
                    ax2.plot(df[c_x], df[c_y2[0]], color=col2, marker='o', linewidth=2, label=c_y2[0])
                    ax2.grid(False)
            elif c_type == "bar":
                for i,c in enumerate(c_y1): ax1.bar(df[c_x], df[c], color=col1 if i==0 else None, alpha=0.8)
            elif c_type == "line":
                for i,c in enumerate(c_y1): ax1.plot(df[c_x], df[c], color=col2 if i==0 else None, marker='o')
            
            ax1.set_title(c_title, pad=15, fontweight='bold')
            ax1.grid(True, linestyle='--', alpha=0.5)
            st.pyplot(fig)
            
            # 导出逻辑
            img = io.BytesIO()
            plt.savefig(img, format='png', dpi=300, bbox_inches='tight')
            img.seek(0)
            
            ppt = io.BytesIO()
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(img, Inches(0.5), Inches(0.5), width=Inches(9))
            prs.save(ppt)
            ppt.seek(0)
            img.seek(0)
            
            d1, d2 = st.columns(2)
            d1.download_button("📥 下载 PNG", img, "chart.png", "image/png", key="dl_1")
            d2.download_button("📥 下载 PPT", ppt, "chart.pptx", key="dl_2")

# ==========================================
# 4. 模块：智能会议纪要 (Q&A版)
# ==========================================
def module_meeting(api_key):
    st.header("🎙️ 智能会议纪要")
    st.markdown("上传录音 -> 自动转写 -> 生成 **Q&A 结构化** 纪要。")
    
    f = st.file_uploader("上传录音", type=['mp3','wav','m4a'], key="mf")
    
    if f and st.button("开始分析", key="btn_meet"):
        if not api_key: st.error("No API Key"); return
        
        st.info("🔄 模拟 ASR 转写中... (真实环境需对接长音频接口)")
        time.sleep(1.5)
        mock_text = """
        王总：Q1营收100亿，增20%。分析师A：毛利为何降？王总：铜价涨了。
        分析师B：未来指引？王总：下半年回升至30%。
        """
        st.text_area("识别结果", mock_text)
        
        dashscope.api_key = api_key
        prompt = f"将以下文本整理为行研会议纪要。要求：1.核心观点。2.Q&A环节严格按'Q: A:'格式。文本：{mock_text}"
        
        try:
            resp = dashscope.Generation.call(model='qwen-turbo', prompt=prompt)
            st.markdown("### 📝 纪要预览")
            st.markdown(resp.output.text)
            st.download_button("下载 TXT", resp.output.text, "minutes.txt", key="dl_txt")
        except: pass

# ==========================================
# 5. 主程序入口
# ==========================================
with st.sidebar:
    st.title("🚀 行研 Copilot")
    api_key = st.text_input("🔑 API Key", type="password", key="mk")
    st.divider()
    mode = st.radio("功能导航", ["🕵️ 全格式核对", "📊 智能制图", "🎙️ 会议纪要"], key="nav")

if mode == "🕵️ 全格式核对": module_compliance(api_key)
elif mode == "📊 智能制图": module_smart_chart_ref(api_key)
elif mode == "🎙️ 会议纪要": module_meeting(api_key)