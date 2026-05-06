from pathlib import Path

from docx import Document


DOC_PATH = Path(r"C:\Coding\WebDetectionSystem-thesis\WebDetectionSystem\基于Python的木质板材缺陷检测系统的设计与实现.docx")


UPDATED_CODE = {
    "图5-9 系统设置页面截图": """# app.py
@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    s = SystemSettings.query.filter_by(key_name='conf_threshold').first()
    s.value = request.form.get('conf_threshold')
    f = request.files.get('model_file')
    if f and f.filename.endswith('.pt'):
        m = SystemSettings.query.filter_by(key_name='current_model').first()
        m.value = secure_filename(f.filename)""",
}


def find_next_nonempty(paragraphs, start_idx):
    """找到占位段后最近的非空段落，作为代码段进行替换。"""
    for idx in range(start_idx + 1, len(paragraphs)):
        if paragraphs[idx].text.strip():
            return paragraphs[idx]
    return None


def main():
    doc = Document(DOC_PATH)
    paragraphs = doc.paragraphs

    for i, para in enumerate(paragraphs):
        caption = para.text.strip()
        if caption not in UPDATED_CODE:
            continue

        for j in range(i + 1, len(paragraphs)):
            if paragraphs[j].text.strip() == "部分关键代码如下：":
                code_para = find_next_nonempty(paragraphs, j)
                if code_para is not None and code_para.text.strip().startswith("#"):
                    code_para.text = UPDATED_CODE[caption]
                break

    doc.save(DOC_PATH)


if __name__ == "__main__":
    main()
