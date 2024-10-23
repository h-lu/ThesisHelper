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

def generate_all_ai_content(task_description, start_date, end_date, title, student_name):
    system_prompt = f"""
    根据以下论文任务书描述，为16次学生论文咨询生成内容。每次咨询包括学生信息和教师信息，每条信息不少于30字且不超过100字，不要有称呼。
    要求每条信息具体详实，包含实质性的内容，避免空泛的表述。
    确保生成的内容与论文任务书相关，并按照论文写作的进度逐步推进。

    论文题目：{title}
    论文任务书描述：
    {task_description}

    咨询时间范围：
    开始日期：{start_date.strftime('%Y-%m-%d')}
    结束日期：{end_date.strftime('%Y-%m-%d')}

    输出格式为JSON，包含以下字段：
    1. consultations: 16个对象的数组，每个对象有date, student_info和teacher_info三个字段。
    2. work_summary: 不超过200字的毕业论文工作总结。这份总结应该是指导教师对学生 {student_name} 的工作评价。
    3. mid_term_review: 不超过100字的中期检查评价，包括指导老师对学生前期工作的评价和对后阶段的要求。简洁扼要。

    示例输出格式：
    {{
        "consultations": [
            {{
                "date": "2024-03-01",
                "student_info": "完成了初步的文献调研，阅读了30余篇相关论文，整理出三个主要研究方向：深度学习在图像识别中的应用、模型优化方法、以及性能评估指标。希望得到老师对研究方向的指导。",
                "teacher_info": "建议聚焦于深度学习在特定场景下的图像识别应用，可以选择医疗影像或工业检测等具体领域。同时要注意总结现有方法的优缺点，为创新点的提出做准备。"
            }},
            // ... 其他14次咨询 ...
            {{
                "date": "2024-06-15",
                "student_info": "完成了所有实验的数据分析和论文的最终修改，对比实验表明新方法在准确率和效率上都有显著提升。已经按照规范要求检查了论文格式，准备好答辩材料。",
                "teacher_info": "论文整体结构完整，实验设计合理，数据分析充分。建议在答辩时重点强调方法创新点和实验结果的可靠性，准备好应对评委可能提出的问题。"
            }}
        ],
        "work_summary": "该生在毕业论文研究过程中表现出色。论文选题具有实际意义，研究方法科学规范。通过大量的文献阅读和实验，提出了创新性的解决方案。工作态度认真，善于思考，能够独立解决问题。论文质量较高，具有一定的学术价值和应用前景。",
        "mid_term_review": "前期工作扎实，文献综述全面，研究方案设计合理。建议在后期工作中加强实验数据的分析深度，突出研究的创新点，注意论文结构的逻辑性和完整性。"
    }}
    """

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "请根据论文任务书和时间范围生成16次论文咨询的内容，包括日期、学生信息和教师信息，以及工作总结和中期检查评价。"}
    ]

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=messages,
        response_format={
            'type': 'json_object'
        }
    )

    return json.loads(response.choices[0].message.content)

def generate_consultations(task_description, start_date, end_date, title, student_name):
    consultations = []

    if 'ai_content' not in st.session_state:
        st.session_state.ai_content = None

    if st.button("使用AI生成所有咨询内容、工作总结和中期检查评价"):
        with st.spinner('正在生成内容...'):
            st.session_state.ai_content = generate_all_ai_content(task_description, start_date, end_date, title, student_name)
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
        location = st.text_input(f"地点", value="办公室", key=f"location_{i}")
        
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

def generate_task_description(title, major, start_date, end_date):
    total_weeks = ((end_date - start_date).days + 1) // 7
    system_prompt = f"""
    请根据给定的论文题目、专业和时间范围，生成一份详细的毕业论文任务书描述。描述应包括以下6个部分：
    1. 课题的任务内容
    2. 原始条件及要求
    3. 设计的技术要求（论文的研究要求）
    4. 毕业设计（论文）应完成的具体工作
    5. 资料文献要求及主要的参考文献
    6. 进度安排（分为4-6个主要阶段，每个阶段说明主要任务和时间安排）

    论文题目：{title}
    专业：{major}
    总周数：{total_weeks}周

    请确保生成的内容专业、详细且与题目和专业相关。每个部分的内容应该简明扼要，总字数控制在1000字左右。
    请以JSON格式输出，每个部分作为一个单独的字段。对于每个字段，如果内容包含多个要点，请使用数组格式，每个要点作为数组的一个元素。

    特别注意：
    - 进度安排应该分为4-6个主要阶段，每个阶段说明主要任务和周数范围。
    - 进度安排应该考虑给定的总周数，确保整个计划在这个时间范围内完成。
    - 使用"x-y周"的格式来表示每个阶段的时间范围，例如"1-3周"。

    示例输出格式：
    {{
        "task_content": [
            "1. 第一个任务内容要点",
            "2. 第二个任务内容要点",
            "3. 第三个任务内容要点"
        ],
        "original_conditions": [
            "1. 第一个原始条件",
            "2. 第二个原始条件"
        ],
        "technical_requirements": [
            "1. 第一个技术要求",
            "2. 第二个技术要求",
            "3. 第三个技术要求"
        ],
        "specific_work": [
            "1. 第一项具体工作",
            "2. 第二项具体工作"
        ],
        "reference_requirements": [
            "1. 第一个参考文献要求",
            "2. 第二个参考文献要求",
            "3. 主要参考文献列表"
        ],
        "schedule": [
            "第一阶段（1-3周）：确定研究主题，完成文献综述",
            "第二阶段（4-8周）：设计研究方法，开始数据收集",
            "第三阶段（9-13周）：完成数据分析，形成初步结果",
            "第四阶段（14-{total_weeks}周）：撰写论文，修改完善，准备答辩"
        ]
    }}
    """

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "请生成毕业论文任务书描述。"}
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

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("执行开始日期", datetime.now())
    with col2:
        end_date = st.date_input("执行结束日期", datetime.now() + timedelta(days=150))

    mid_date = start_date + (end_date - start_date) / 2

    student_signature_file = st.file_uploader("上传学生签名图片", type=["png", "jpg", "jpeg"])
    teacher_signature_file = st.file_uploader("上传教师签名图片", type=["png", "jpg", "jpeg"])

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
                task_content = generate_task_description(title, major, start_date, end_date)
            
            # 更新 session_state 中的 task_parts
            st.session_state.task_parts = task_content
        
        # 显示生成的内容并允许编辑
        for i, (key, part_name) in enumerate([
            ("task_content", "课题的任务内容"),
            ("original_conditions", "原始条件及要求"),
            ("technical_requirements", "设计的技术要求"),
            ("specific_work", "应完成的具体工作"),
            ("reference_requirements", "资料文献要求"),
            ("schedule", "进度安排")
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
                    
                    # 准备渲染上下文
                    task_context = {
                        'title': title,
                        'student_name': student_name,
                        'student_id': student_id,
                        'teacher_name': teacher_name,
                        'teacher_signature': teacher_signature,
                        'major': major,
                        'college': college,
                        'start_date': start_date.strftime("%Y-%m-%d"),
                        'end_date': end_date.strftime("%Y-%m-%d"),
                        **st.session_state.task_parts  # 使用 session_state 中的 task_parts
                    }
                    
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

        if student_signature_file and teacher_signature_file and st.session_state.task_parts:
            # 使用任务书内容生成咨询记录
            task_description = "\n".join(["\n".join(st.session_state.task_parts[key]) for key in st.session_state.task_parts])
            
            # 生成咨询记录、工作总结和中期检查评价
            consultations, work_summary, mid_term_review = generate_consultations(task_description, start_date, end_date, title, student_name)

            if st.button("生成咨询记录"):
                # 加载签名图片
                student_signature = InlineImage(doc, student_signature_file, width=Mm(20))
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
                    'student_signature': student_signature,
                    'teacher_signature': teacher_signature,
                    'work_summary': work_summary,
                    'mid_term_review': mid_term_review,
                    'pagebreak': RichText('\f')
                }
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
            st.warning("请先生成任务书内容，并上传签名图片。")

if __name__ == "__main__":
    main()

