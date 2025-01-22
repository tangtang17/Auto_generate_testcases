import openai
import openpyxl
import docx

# 设置 OpenAI API 密钥
openai.api_key = "openai_api_key"

# 读取需求内容
def read_requirements_from_document(document_path):
    # 假设需求文档是 Word 格式
    doc = docx.Document(document_path)
    text = []
    for para in doc.paragraphs:
        if para.text.strip():
            # 提取非空字段
            text.append(para.text.strip())
    return text

# 从需求文档生成测试用例的函数
def generate_test_cases_from_requirements(requirements):
    test_cases = []

    for req in requirements:
        # 向 GPT-3 请求生成测试用例
        response = openai.Completion.create(
                model="gpt-3.5",  # or gpt-3.5
                prompt=f"""
                根据以下需求生成测试用例，包含测试用例描述、步骤和预期结果：\n
                需求：{req}\n\n
                生成的测试用例应包括：
                - Test Case ID
                - Test Description
                - Test Steps
                - Expected Results
                - Status (Pass, Fail, Blocked)\n
                \n测试用例：
            """,
                max_tokens=300,
                temperature=0.7
            )

        test_case = response.choices[0].text.strip()
        # 状态默认为空
        test_cases.append({"Requirement": req, "Test Case": test_case, "Status": ""})

    return test_cases

# 写入测试用例到 Excel 文件
def write_test_cases_to_excel(test_cases, output_file="test_cases.xlsx"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Test Cases"

    # 写入表头
    headers = ["Requirement", "Test Case Description", "Test Steps", "Expected Results", "Status"]
    sheet.append(headers)

    # 写入每个测试用例
    for case in test_cases:
        test_case = case["Test Case"].split("\n")
        test_description = test_case[0] if len(test_case) > 0 else ""
        test_steps = test_case[1] if len(test_case) > 1 else ""
        expected_results = test_case[2] if len(test_case) > 2 else ""
        status = case["Status"]
        row = [case["Requirement"], test_description, test_steps, expected_results, status]
        sheet.append(row)

    # 保存文件
    workbook.save(output_file)
    print(f"Test cases saved to {output_file}")


def main():
    file_path = "requirements.docx"
    requirements = read_requirements_from_document(file_path)

    # 生成测试用例
    test_cases = generate_test_cases_from_requirements(requirements)

    # 保存为 Excel 文件
    write_test_cases_to_excel(test_cases)

if __name__ == "__main__":
    main()

