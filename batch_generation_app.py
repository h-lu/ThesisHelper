import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText, InlineImage
from datetime import datetime, timedelta
from docx.shared import Mm
import io
import base64
import os
import zipfile
import tempfile
from student_consultation_app import (
    generate_task_description,
    generate_all_ai_content,
    consultation_keywords
)

def extract_signatures(zip_file):
    """解压签名文件到临时目录"""
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_file) as zf:
        # 获取文件列表
        file_list = zf.namelist()
        
        # 过滤掉 macOS 系统文件
        valid_files = [f for f in file_list if not f.startswith('__MACOSX')]
        
        # 解压文件
        for file in valid_files:
            try:
                # 使用UTF-8解码文件名
                decoded_name = file.encode('cp437').decode('utf-8')
                # 提取文件
                with zf.open(file) as source, open(os.path.join(temp_dir, decoded_name), 'wb') as target:
                    target.write(source.read())
            except Exception as e:
                st.error(f"处理文件时出错：{str(e)}")
    
    return temp_dir

def process_excel_file(excel_file):
    """读取并处理CSV文件"""
    try:
        df = pd.read_csv(excel_file)
        required_columns = [
            "论文题目", "学生姓名", "学生学号", "指导教师", 
            "专业", "学院", "开始日期", "结束日期"
        ]
        
        # 检查必要的列是否存在
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"CSV文件缺少以下列：{', '.join(missing_columns)}")
            return None
            
        return df
    except Exception as e:
        st.error(f"处理CSV文件时出错：{str(e)}")
        return None

def generate_documents_for_student(row, teacher_signature_file, dean_signature_file, signatures_dir):
    """为单个学生生成文档"""
    try:
        # 转换日期格式
        start_date = pd.to_datetime(row["开始日期"]).date()
        end_date = pd.to_datetime(row["结束日期"]).date()
        
        # 获取学生签名图片路径（使用学生姓名作为文件名）
        # 尝试多个可能的扩展名
        student_signature_path = None
        for ext in ['.jpg', '.jpeg', '.png']:
            path = os.path.join(signatures_dir, f"{row['学生姓名']}{ext}")
            if os.path.exists(path):
                student_signature_path = path
                break
                
        if not student_signature_path:
            st.error(f"找不到学生 {row['学生姓名']} 的签名图片")
            return None, None
        
        # 生成任务书内容
        task_content = generate_task_description(
            row["论文题目"], 
            row["专业"], 
            start_date, 
            end_date
        )
        
        if not task_content:
            return None, None
            
        # 将列表转换为多行文本
        formatted_task_content = {}
        for key in task_content:
            if isinstance(task_content[key], list):
                formatted_task_content[key] = '\n'.join(task_content[key])
            else:
                formatted_task_content[key] = task_content[key]
            
        # 生成咨询记录内容
        task_description = "\n".join([formatted_task_content[key] for key in formatted_task_content])
        ai_content = generate_all_ai_content(
            task_description,
            start_date,
            end_date,
            row["论文题目"],
            row["学生姓名"]
        )
        
        # 生成任务书文档
        task_doc = DocxTemplate("thesis_task_description_template.docx")
        teacher_signature = InlineImage(task_doc, teacher_signature_file, width=Mm(20))
        student_signature = InlineImage(task_doc, student_signature_path, width=Mm(20))
        dean_signature = InlineImage(task_doc, dean_signature_file, width=Mm(20))
        
        task_context = {
            'title': row["论文题目"],
            'student_name': row["学生姓名"],
            'student_id': row["学生学号"],
            'teacher_name': row["指导教师"],
            'teacher_signature': teacher_signature,
            'student_signature': student_signature,
            'dean_signature': dean_signature,
            'major': row["专业"],
            'college': row["学院"],
            'start_date': start_date.strftime("%Y-%m-%d"),
            'end_date': end_date.strftime("%Y-%m-%d"),
            **formatted_task_content  # 使用格式化后的内容
        }
        
        task_doc.render(task_context)
        
        # 生成记录本文档
        record_doc = DocxTemplate("student_consultation_template.docx")
        teacher_signature = InlineImage(record_doc, teacher_signature_file, width=Mm(20))
        student_signature = InlineImage(record_doc, student_signature_path, width=Mm(20))
        
        mid_date = start_date + (end_date - start_date) / 2
        
        record_context = {
            'title': row["论文题目"],
            'student_name': row["学生姓名"],
            'student_id': row["学生学号"],
            'teacher_name': row["指导教师"],
            'teacher_signature': teacher_signature,
            'student_signature': student_signature,
            'dean_signature': dean_signature,
            'major': row["专业"],
            'college': row["学院"],
            'start_date': start_date.strftime("%Y-%m-%d"),
            'mid_date': mid_date.strftime("%Y-%m-%d"),
            'end_date': end_date.strftime("%Y-%m-%d"),
            'consultations': ai_content['consultations'],
            'work_summary': ai_content['work_summary'],
            'mid_term_review': ai_content['mid_term_review'],
            'pagebreak': RichText('\f')
        }
        
        record_doc.render(record_context)
        
        return task_doc, record_doc
        
    except Exception as e:
        st.error(f"生成文档时出错：{str(e)}")
        st.exception(e)  # 显示详细的错误信息
        return None, None

def get_csv_download_link():
    """生成CSV模板文件的下载链接"""
    with open("template.csv", "r", encoding='utf-8') as f:
        csv_content = f.read()
    b64 = base64.b64encode(csv_content.encode()).decode()
    href = f'<a href="data:text/csv;base64,{b64}" download="template.csv">模板文件</a>'
    return href

def main():
    st.title("批量生成毕业论文归档材料")
    
    # 添加使用说明
    with st.expander("使用说明（请先阅读）", expanded=True):
        st.markdown("""
        ### 准备工作
        
        1. **准备CSV文件**
           - 下载""")
        st.markdown(get_csv_download_link(), unsafe_allow_html=True)
        st.markdown("""
           - 按模板格式填写学生信息
           - 必须包含以下列：论文题目、学生姓名、学生学号、指导教师、专业、学院、开始日期、结束日期
           - 日期格式为：YYYY-MM-DD（如：2024-03-01）
        
        2. **准备签名图片**
           - 教师签名：单个图片文件（支持jpg、jpeg、png格式）
           - 学生签名：
             * 每个学生一个签名图片文件
             * 文件名必须与CSV中的"学生姓名"完全一致（如：张三.jpg）
             * 支持jpg、jpeg、png格式
             * 将所有学生签名图片打包成一个ZIP文件
        
        ### 使用步骤
        
        1. 上传CSV文件
        2. 上传教师签��图片
        3. 上传包含所有学生签名ZIP文件
        4. 点击"开始批量生成文档"
        5. 等待处理完成后下载生成的ZIP文件
        
        ### 注意事项
        
        - 签名图片建议使用白色背景
        - 签名图片大小建议不超过1MB
        - ZIP文件中不要包含文件夹，直接放签名图片
        - 确保所有文件名中不包含特殊字符
        - 生成的文档将按"学生姓名 - 任务书.docx"和"学生姓名 - 记录本.docx"的格式命名
        
        ### 输出文件
        
        程序将生成一个ZIP文件，包含：
        1. 每个学生的任务书（包含任务内容、进度安排等）
        2. 每个学生的记录本（包含16次咨询记录、中期检查评价和工作总结）
        """)
    
    # 传CSV文件
    excel_file = st.file_uploader("上传学生信息CSV文件", type=["csv"])
    
    # 上传教师签名
    teacher_signature_file = st.file_uploader("上传教师签名图片", type=["png", "jpg", "jpeg"])

    # 上传系主任签名
    dean_signature_file = st.file_uploader("上传系主任签名图片", type=["png", "jpg", "jpeg"])
    
    # 上传学生签名ZIP文件
    signatures_zip = st.file_uploader("上传学生签名ZIP文件（签名图片文件名需与学生姓名一致）", type="zip")
    
    if excel_file and teacher_signature_file and dean_signature_file and signatures_zip:  # 添加dean_signature_file检查
        # 解压签名文件到临时目录
        signatures_dir = extract_signatures(signatures_zip)
        # st.write(f"临时目录路径: {signatures_dir}")
            
        df = process_excel_file(excel_file)
        
        if df is not None:
            st.write("已读取的学生信息：")
            st.dataframe(df)
            
            # 验证所有学生的签名图片是否存在
            missing_signatures = []
            for _, row in df.iterrows():
                # 检查是否存在任一格式的签名文件
                found = False
                for ext in ['.jpg', '.jpeg', '.png']:
                    file_path = os.path.join(signatures_dir, f"{row['学生姓名']}{ext}")
                    # st.write(f"检查文件: {file_path}")
                    if os.path.exists(file_path):
                        found = True
                        # st.write(f"找到签名文件: {file_path}")
                        break
                if not found:
                    missing_signatures.append(row['学生姓名'])
            
            if missing_signatures:
                st.error(f"以下学生的签名图片未找到：{', '.join(missing_signatures)}")
                return
            
            if st.button("开始批量生成文档"):
                # 创建ZIP文件
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for index, row in df.iterrows():
                        with st.spinner(f"正在处理 {row['学生姓名']} 的文档..."):
                            task_doc, record_doc = generate_documents_for_student(
                                row, 
                                teacher_signature_file,
                                dean_signature_file,  # 添加系主任签名参数
                                signatures_dir
                            )
                            
                            if task_doc and record_doc:
                                # 保存任务书
                                task_buffer = io.BytesIO()
                                task_doc.save(task_buffer)
                                zf.writestr(
                                    f"{row['学生姓名']} - 任务书.docx",
                                    task_buffer.getvalue()
                                )
                                
                                # 保存记录本
                                record_buffer = io.BytesIO()
                                record_doc.save(record_buffer)
                                zf.writestr(
                                    f"{row['学生姓名']} - 记录本.docx",
                                    record_buffer.getvalue()
                                )
                
                # 提供ZIP文件下载
                zip_buffer.seek(0)
                b64 = base64.b64encode(zip_buffer.getvalue()).decode()
                href = f'<a href="data:application/zip;base64,{b64}" download="毕业论文归档材料.zip">下载所有生成的文档</a>'
                st.markdown(href, unsafe_allow_html=True)
                
                st.success("所有文档已生成完成！")
                
                # 清理临时目录
                import shutil
                shutil.rmtree(signatures_dir)

if __name__ == "__main__":
    main()
