import streamlit as st
from docxtpl import DocxTemplate, RichText, InlineImage
from datetime import datetime, timedelta
from docx.shared import Mm
import io
import base64
import json
from openai import OpenAI

# 初始化OpenAI客户端
client = OpenAI(
    api_key=st.secrets["DEEPSEEK_API_KEY"],
    base_url="https://api.deepseek.com",
)

# 定义每次咨询的关键字
consultation_keywords = [
    {"student": "选题讨论", "teacher": "方向建议"},
    {"student": "文献综述", "teacher": "资料推荐"},
    {"student": "研究方法", "teacher": "实验设计"},
    {"student": "数据收集", "teacher": "分析方法"},
    {"student": "初步结果", "teacher": "改进建议"},
    {"student": "论文大纲", "teacher": "结构优化"},
    {"student": "实验进展", "teacher": "数据解释"},
    {"student": "章节撰写", "teacher": "内容审阅"},
    {"student": "统计分析", "teacher": "结果讨论"},
    {"student": "图表制作", "teacher": "可视化建议"},
    {"student": "讨论部分", "teacher": "深度分析"},
    {"student": "结论总结", "teacher": "贡献点确认"},
    {"student": "摘要撰写", "teacher": "关键词确定"},
    {"student": "参考文献", "teacher": "格式检查"},
    {"student": "论文定稿", "teacher": "最终修改"},
    {"student": "答辩准备", "teacher": "预答辩指导"}
]

def generate_all_ai_content(task_description, start_date, end_date, title, student_name, additional_info=""):
    system_prompt = f"""
    根据以下论文任务书描述和补充信息，为16次学生论文咨询生成内容。每次咨询包括学生信息和教师信息，具体要求如下：

    基本要求：
    1. 每条信息100-200字，确保内容充实且有实质性指导价值
    2. 每条信息包含3-4个完整的句子
    3. 内容具体详实，避免空泛表述，需包含具体的研究细节、方法和建议
    4. 不要有称呼语，直接描述内容
    5. 按照论文写作的进度逐步推进，体现研究的连续性和深入性
    6. 每次咨询都要体现实质性进展，不能简单重复

    内容要求：
    1. 学生信息应包含：
       - 当前工作的具体进展
       - 遇到的具体问题或困难
       - 已经采取的解决方案
       - 下一步的工作计划

    2. 教师信息应包含：
       - 对学生工作的具体评价
       - 针对性的改进建议
       - 明确的指导方向
       - 具体的技术或方法建议

    3. 进度安排：
       - 前5次咨询：选题定位、文献研究、方法设计阶段
       - 中5次咨询：实验/调研实施、数据收集分析阶段
       - 后6次咨询：论文撰写、修改完善阶段

    论文信息：
    论文题目：{title}
    论文任务书描述：{task_description}
    补充信息：{additional_info}

    时间安排：
    开始日期：{start_date.strftime('%Y-%m-%d')}
    结束日期：{end_date.strftime('%Y-%m-%d')}

    输出格式为JSON，包含以下字段：
    1. consultations: 16个对象的数组，每个对象包含：
       - date: 咨询日期
       - student_info: 学生工作汇报（100-200字）
       - teacher_info: 教师指导建议（100-200字）

    2. work_summary: 200-300字的毕业论文工作总结，包含：
       - 总结学生的工作态度和表现
       - 评价研究工作的创新性和价值
       - 对论文质量的整体评价
       - 对学生的期望和建议

    3. mid_term_review: 150-200字的中期检查评价，包含：
       - 前期工作的具体评价
       - 已取得的阶段性成果
       - 存在的问题和不足
       - 后工作的具体要求和建议

    示例输出格式：
    {{
        "consultations": [
            {{
                "date": "2024-03-01",
                "student_info": "完成了20篇核心期刊论文的系统阅读和分析，重点关注了深度学习在图像识别领域的最新进展。通过文献梳理，发现目前主要存在模型复杂度高和泛化能力不足两个问题。基于文献分析结果，初步构思了一个基于轻量级网络的改进方案，并完成了技术路线的初步设计。准备开始进行算法的详细设计和实验环境的搭建。",
                "teacher_info": "文献综述工作比较系统，问题定位准确。建议进一步细化改进方案中的创新点，可以从模型结构优化和损失函数设计两个方向深入。同时要注意收集足够的实验数据，建议准备至少三个公开数据集进行验证。需要设计详细的对比实验方案，确保研究结果的可靠性和说服力。"
            }}
        ],
        "work_summary": "该生在毕业论文研究过程中表现出色，工作态度认真负责，科研能力突出。论文选题紧跟学科前沿，具有重要的理论意义和应用价值。在研究过程中，通过大量的文献阅读和实验探索，提出了具有创新性的解决方案。实验设计严谨，数据分析深入，研究结果可靠。特别值得肯定的是，该生善于思考，能够独立解决问题，具备良好的科研素养。论文质量较高，创新点明确，实验验证充分，具有较好的学术价值和应用前景。",
        "mid_term_review": "前期工作扎实，文献综述全面且深入，研究方案设计合理可行。已完成关键算法的设计和初步实验，取得了积极的阶段性成果。存在的问题是实验验证还需要进一步深入，数据分析有待加强。建议在后期工作中重点加强实验数据的分析深度，进一步突出研究的创新点，同时注意论文结构的逻辑性和完整性。要按计划推进实验工作，确保留出充足的论文修改时间。"
    }}
    """

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "请根据给定的论文题目和专业生成一个JSON格式的任务书描述。"}
    ]

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=messages,
        response_format={
            'type': 'json_object'
        }
    )

    return json.loads(response.choices[0].message.content)

def generate_consultations(task_description, start_date, end_date, title, student_name, additional_info=""):
    consultations = []

    if 'ai_content' not in st.session_state:
        st.session_state.ai_content = None

    if st.button("使用AI生成所有咨询内容、工作总结和中期检查评价"):
        with st.spinner('正在生成内容...'):
            st.session_state.ai_content = generate_all_ai_content(task_description, start_date, end_date, title, student_name, additional_info)
        st.success("所有内容已生成!")

    # 验证AI生成的内容
    if st.session_state.ai_content and 'consultations' in st.session_state.ai_content:
        ai_consultations = st.session_state.ai_content['consultations']
        if len(ai_consultations) != 16:
            st.warning(f"AI生成的咨询记录数量不正确。预期16条，实际生成{len(ai_consultations)}条。将使用默认值填充。")
    else:
        ai_consultations = []

    for i in range(16):
        st.text(f"咨询 {i+1}")
        
        if i < len(ai_consultations):
            if i == 0:
                default_date = start_date.strftime("%Y-%m-%d")
            elif i == 15:
                default_date = end_date.strftime("%Y-%m-%d")
            else:
                default_date = ai_consultations[i].get('date', (start_date + (end_date - start_date) * i / 15).strftime("%Y-%m-%d"))
            default_student_info = ai_consultations[i].get('student_info', "")
            default_teacher_info = ai_consultations[i].get('teacher_info', "")
        else:
            if i == 0:
                default_date = start_date.strftime("%Y-%m-%d")
            elif i == 15:
                default_date = end_date.strftime("%Y-%m-%d")
            else:
                default_date = (start_date + (end_date - start_date) * i / 15).strftime("%Y-%m-%d")
            default_student_info = ""
            default_teacher_info = ""
        
        time = st.text_input(f"时间", value=default_date, key=f"time_{i}")
        location = st.text_input(f"地点", value="办公", key=f"location_{i}")
        
        student_input = st.text_area(f"学生信息 (关键词: {consultation_keywords[i]['student']})", 
                                     value=default_student_info,
                                     key=f"student_input_{i}")
        teacher_input = st.text_area(f"教师信息 (关键词: {consultation_keywords[i]['teacher']})", 
                                     value=default_teacher_info,
                                     key=f"teacher_input_{i}")
        
        consultation = {
            'id': i + 1,
            'time': time,
            'location': location,
            'student_info': student_input,
            'teacher_info': teacher_input
        }
        consultations.append(consultation)
    


    # 显示中期检查评价
    if st.session_state.ai_content and 'mid_term_review' in st.session_state.ai_content:
        mid_term_review = st.text_area("中期检查评价（可编辑）", value=st.session_state.ai_content['mid_term_review'], height=200)
    else:
        mid_term_review = st.text_area("中期检查评价（可编辑）", value="", height=200)

     # 显示工作总结
    if st.session_state.ai_content and 'work_summary' in st.session_state.ai_content:
        work_summary = st.text_area("工作总结（可编辑）", value=st.session_state.ai_content['work_summary'], height=300)
    else:
        work_summary = st.text_area("工作总结（可编辑）", value="", height=300)
    
    return consultations, work_summary, mid_term_review

def generate_task_description(title, major, start_date, end_date, additional_info=""):
    total_weeks = ((end_date - start_date).days + 1) // 7
    system_prompt = f"""
    请根据给定的论文题目、专业、时间范围和补充信息，生成一份详细的毕业论文任务书描述。描述应包括以下5个部分，并以JSON格式输出：

    论文题目：{title}
    专业：{major}
    补充信息：{additional_info}

    1. 课题的任务内容：
       - 须融入论文选题内容，不少于100字
       - 阐述选题的研究背景和现实意义
       - 明确研究目标和预期成果
       - 说明研究的创新点和应用价值
       - 指出研究的重点和难点
       - 说明研究的理论和实践意义
       - 阐述研究的可行性分析

    2. 原始条件及数据：
       - 说明完成论文所需的基础知识和技能要求
       - 列出必要的软硬件环境和工具
       - 明确数据来源和获取方式
       - 说明数据的类型和规模
       - 规定数据的质量要求
       - 说明数据的预处理方法
       - 规定数据的存储和管理方式

    3. 设计的技术要求（论文的研究要求）：
       - 详细说明研究方法和技术路线
       - 提出具体的技术指标和参数要求
       - 规定实验或调研的具体要求
       - 明确数据处理和分析方法
       - 提出创新性要求和技术突破点
       - 说明研究的可验证性
       - 规定研究结果的评价标准

    4. 毕业设计（论文）应完成的具体工作：
       A. 基本要求（通用部分）：
          1. 文献综述和开题报告：
             - 开题报告成绩要求70分以上合格
             - 文献综述字数2500字左右
             - 开题报告需包含研究计划和预期目标
          2. 外文翻译：
             - 翻译一篇与选题相关的英文文献
             - 字数要求20000英文印刷字符以上
             - 翻译质量要准确、通顺
          3. 调研工作：
             - 进行实地调研或实验研究
             - 调研报告字数3000字左右
             - 需包含数据分析和结果讨论
          4. 论文撰写：
             - 论文总字数1.5~2万字
             - 符合学校论文格式规范
             - 完成导师要求的修改
          5. 论文答辩：
             - 准备答辩PPT和讲稿
             - 参加答辩并回答问题
             - 总分60分以上为通过

       B. 研究工作（根据论文题目"{title}"和专业"{major}"生成具体内容）：
          1. 理论研究部分：
             - 系统梳理本研究领域的理论基础
             - 构建适合研究问题的理论框架
             - 提出研究假设或理论模型
             - 确定关键变量和影响因素
          2. 研究方法部分：
             - 设计详细的研究方案
             - 确定研究方法和技术路线
             - 制定数据收集和分析计划
             - 建立评估指标体系
          3. 实验/调研部分：
             - 开展实验或调研工作
             - 收集和整理原始数据
             - 进行数据预处理和分析
             - 验证研究假设
          4. 创新工作部分：
             - 提出创新性的解决方案
             - 设计和实施对比实验
             - 总结研究的创新点
             - 验证创新成果的有效性
          5. 应用研究部分：
             - 选择典型案例进行分析
             - 进行实践应用验证
             - 评估应用效果
             - 总结实践价值和推广意义

    5. 资料文献要求及主要的参考文献：
       - 文献数量要求：
         * 外文文献不少于4篇
         * 中文文献不少于16篇
         * 核心期刊文献占比不低于50%
       - 文献时效性要求：
         * 近五年文献占比不少于50%
         * 需包含最新研究进展
       - 文献搜索途径：
         * 外文数据库：Web of Science、Scopus、IEEE Xplore等
         * 中文数据库：CNKI、万方、维普等
         * 学术搜索引擎：Google Scholar、百度学术等
       - 文献类型要求：
         * 以学术期刊论文为主
         * 必须包含核心期刊文献
         * 可包含高水平会议论文
         * 可包含优秀博硕士论文
       - 文献引用规范：
         * 遵守学术规范
         * 注意避免过度引用
         * 引用格式符合要求
       - 建议关键词：根据论文主题提供5-8个核心关键词
       - 推荐经典文献：列出3-5篇该领域的经典或高被引文献

    请确保生成的内容：
    1. 专业性：使用专业术语和表达方式
    2. 针对性：内容与论文题目和专业紧密相关
    3. 可操作性：要求具体明确，便于执行
    4. 完整性：覆盖论文写作的各个环节
    5. 规范性：符合学术规范和学校要求
    6. 总字数：控制在1000字左右

    请生成一个JSON格式的输出，每个部分作为一个单独的字段。对于每个字段，如果内容包含多个要点，请使用数组格式，每个要点作为数组的一个元素。

    输出的JSON格式示例：
    {{
        "task_content": [
            "1. 研究背景：...(详细阐述选题背景和意义，不少于100字)",
            "2. 研究目标：...(明确具体的研究目标)",
            "3. 创新点：...(说明研究的创新之处)",
            "4. 研究重点和难点：...(指出关键问题)",
            "5. 理论和实践意义：...(阐述研究价值)",
            "6. 可行性分析：...(说明研究的可行性)"
        ],
        "original_conditions": [
            "1. 基础知识要求：...(列出必备知识)",
            "2. 环境和工具要求：...(说明所需环境)",
            "3. 数据来源：...(明确数据来源)",
            "4. 数据类型：...(说明数据类型)",
            "5. 数据规模：...(规定数据规模)",
            "6. 数据质量：...(说明质量要求)",
            "7. 数据管理：...(规定管理方式)"
        ],
        "technical_requirements": [
            "1. 研究方法：...(详述研究方法)",
            "2. 技术指标：...(列出具体指标)",
            "3. 实验要求：...(说明实验规范)",
            "4. 数据分析方法：...(规定分析方法)",
            "5. 创新性要求：...(提出创新要求)",
            "6. 可验证性：...(说明验证方法)",
            "7. 评价标准：...(规定评价标准)"
        ],
        "specific_work": [
            "A. 基本要求（通用部分）：",
            "1. 文献综述和开题报告：",
            "   - 开题报告成绩要求70分以上合格",
            "   - 文献综述字数2500字左右",
            "   - 开题报告需包含研究计划和预期目标",
            "2. 外文翻译：",
            "   - 翻译一篇与选题相关的英文文献",
            "   - 字数要求20000英文印刷字符以上",
            "   - 翻译质量要准确、通顺",
            "3. 调研工作：",
            "   - 进行实地调研或实验研究",
            "   - 调研报告字数3000字左右",
            "   - 需包含数据分析和结果讨论",
            "4. 论文撰写：",
            "   - 论文总字数1.5~2万字",
            "   - 符合学校论文格式规范",
            "   - 完成导师要求的修改",
            "5. 论文答辩：",
            "   - 准备答辩PPT和讲稿",
            "   - 参加答辩并回答问题",
            "   - 总分60分以上为通过",
            "",
            "B. 研究工作（具体内容）：",
            "1. 理论研究：[根据论文题目生成具体的理论研究任务]",
            "2. 研究方法：[根据论文题目生成具体的研究方法]",
            "3. 实验/调研：[根据论文题目生成具体的实验或调研任务]",
            "4. 创新工作：[根据论文题目生成具体的创新任务]",
            "5. 应用研究：[根据论文题目生成具体的应用研究任务]"
        ],
        "reference_requirements": [
            "1. 文献数量和类型要求：",
            "   - 外文文献不少于4篇",
            "   - 中文文献不少于16篇",
            "   - 核心期刊文献占比不低于50%",
            "2. 文献时效性要求：",
            "   - 近五年文献占比不少于50%",
            "   - 需包含最新研究进展",
            "3. 文献搜索途径：",
            "   - 外文数据库：Web of Science、Scopus、IEEE Xplore等",
            "   - 中文数据库：CNKI、万方、维普等",
            "   - 学术搜索引擎：Google Scholar、百度学术等",
            "4. 文献引用规范：",
            "   - 遵守学术规范",
            "   - 注意避免过度引用",
            "   - 引用格式符合要求",
            "5. 建议关键词：[与论文主题相关的5-8个关键词]",
            "6. 推荐经典文献：[3-5篇该领域的经典或高被引文献]"
        ]
    }}
    """

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "请根据给定的论文题目和专业生成一个JSON格式的任务书描述。"}
    ]

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=messages,
        response_format={
            'type': 'json_object'
        }
    )

    return json.loads(response.choices[0].message.content)

def main():
    st.title("毕业论文归档材料生成器")

    # 主页面用户输入
    title = st.text_input("论文题目")
    student_name = st.text_input("学生姓名")
    student_id = st.text_input("学生学号")
    teacher_name = st.text_input("指导教师")
    major = st.text_input("专业")
    college = st.text_input("学院", value="经济与管理学院")

    # 添加额外信息输入框
    additional_info = st.text_area(
        "补充信息（可选）",
        help="请输入一些重要的补充信息，如研究方向的具体要求、特殊的技术路线、导师的具体要求等。这些信息将用于生成更准确的任务书内容。",
        height=100
    )

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("执行开始日期", datetime.now())
    with col2:
        end_date = st.date_input("执行结束日期", datetime.now() + timedelta(days=150))

    mid_date = start_date + (end_date - start_date) / 2

    student_signature_file = st.file_uploader("上传学生签名图片（可选）", type=["png", "jpg", "jpeg"])
    teacher_signature_file = st.file_uploader("上传教师签名图片（必需）", type=["png", "jpg", "jpeg"])
    dean_signature_file = st.file_uploader("上传系主任签名图片（必需）", type=["png", "jpg", "jpeg"])

    # 创建选项卡，"任务书"在前
    tab1, tab2 = st.tabs(["任务书", "记录本"])

    with tab1:
        st.subheader("毕业论文任务书生成")
        
        # 初始化 task_parts
        if 'task_parts' not in st.session_state:
            st.session_state.task_parts = {
                "task_content": [],
                "original_conditions": [],
                "technical_requirements": [],
                "specific_work": [],
                "reference_requirements": [],
                "schedule": []
            }
        
        if st.button("生成任务书内容"):
            with st.spinner("正在生成任务书内容..."):
                task_content = generate_task_description(title, major, start_date, end_date, additional_info)
            
            # 更新 session_state 中的 task_parts
            st.session_state.task_parts = task_content
        
        # 显示生成的内容并允许编辑
        for i, (key, part_name) in enumerate([
            ("task_content", "课题的任务内容"),
            ("original_conditions", "原始条件及数据"),
            ("technical_requirements", "设计的技术要求"),
            ("specific_work", "应完成的具体工作"),
            ("reference_requirements", "资料文献要求")
        ]):
            content = st.session_state.task_parts.get(key, [])
            if isinstance(content, list):
                formatted_content = "\n".join(content)
            else:
                formatted_content = content
            
            st.session_state.task_parts[key] = st.text_area(
                f"{i+1}. {part_name}", 
                value=formatted_content, 
                height=150,
                key=f"task_part_{key}"
            )

        if st.button("生成任务书文档"):
            if not teacher_signature_file:
                st.warning("请上传教师的签名图片。")
            elif not dean_signature_file:
                st.warning("请上传系主任的签名图片。")
            elif not all([title, student_name, student_id, teacher_name, major, college]):
                st.warning("请填写所有基本信息。")
            elif not all(st.session_state.task_parts.values()):
                st.warning("请先生成或填写所有任务书内容。")
            else:
                try:
                    # 加载任务书模板
                    task_doc = DocxTemplate("thesis_task_description_template.docx")
                    
                    # 在这里创建 InlineImage 对象
                    teacher_signature = InlineImage(task_doc, teacher_signature_file, width=Mm(20))
                    dean_signature = InlineImage(task_doc, dean_signature_file, width=Mm(20))
                    
                    # 准备渲染上下文
                    task_context = {
                        'title': title,
                        'student_name': student_name,
                        'student_id': student_id,
                        'teacher_name': teacher_name,
                        'teacher_signature': teacher_signature,
                        'dean_signature': dean_signature,
                        'major': major,
                        'college': college,
                        'start_date': start_date.strftime("%Y-%m-%d"),
                        'end_date': end_date.strftime("%Y-%m-%d"),
                        **st.session_state.task_parts  # 使用 session_state 中的 task_parts
                    }
                    
                    # 如果有学生签名，添加到上下文中
                    if student_signature_file:
                        task_context['student_signature'] = InlineImage(task_doc, student_signature_file, width=Mm(20))
                    
                    # 渲染模板
                    task_doc.render(task_context)
                    
                    # 保存生成的文档到内存中
                    output = io.BytesIO()
                    task_doc.save(output)
                    output.seek(0)

                    # 提供下载链接
                    b64 = base64.b64encode(output.getvalue()).decode()
                    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{student_name} - 任务书.docx">下载生成的任务书</a>'
                    st.markdown(href, unsafe_allow_html=True)

                    st.success(f"{student_name}的毕业论文任务书已生成!")
                except Exception as e:
                    st.error(f"生成文档时出错：{str(e)}")
                    st.error("错误详情：")
                    st.exception(e)

    with tab2:
        # 加载模板
        doc = DocxTemplate("student_consultation_template.docx")

        if teacher_signature_file and st.session_state.task_parts:
            # 使用任务书内容生成咨询记录
            task_description = "\n".join(["\n".join(st.session_state.task_parts[key]) for key in st.session_state.task_parts])
            
            # 生成咨询记录、工作总结和中期检查评价
            consultations, work_summary, mid_term_review = generate_consultations(task_description, start_date, end_date, title, student_name, additional_info)

            if st.button("生成咨询记录"):
                # 加载签名图片
                teacher_signature = InlineImage(doc, teacher_signature_file, width=Mm(20))

                # 渲染模板
                context = {
                    'title': title,
                    'start_date': start_date.strftime("%Y-%m-%d"),
                    'mid_date': mid_date.strftime("%Y-%m-%d"),
                    'end_date': end_date.strftime("%Y-%m-%d"),
                    'student_name': student_name,
                    'student_id': student_id,
                    'teacher_name': teacher_name,
                    'major': major,
                    'college': college,
                    'consultations': consultations,
                    'teacher_signature': teacher_signature,
                    'work_summary': work_summary,
                    'mid_term_review': mid_term_review,
                    'pagebreak': RichText('\f')
                }

                # 如果有学生签名，添加到上下文中
                if student_signature_file:
                    context['student_signature'] = InlineImage(doc, student_signature_file, width=Mm(20))

                doc.render(context)

                # 保存生成的文档到内存中
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)

                # 提供下载链接
                b64 = base64.b64encode(output.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{student_name} - 记录本.docx">下载生成的咨询记录</a>'
                st.markdown(href, unsafe_allow_html=True)

                st.success(f"{student_name}的学生咨询记录已生成!")
        else:
            if not teacher_signature_file:
                st.warning("请上传教师签名图片。")
            if not st.session_state.task_parts:
                st.warning("请先生成任务书内容。")

if __name__ == "__main__":
    main()

