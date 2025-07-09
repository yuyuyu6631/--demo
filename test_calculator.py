# 导入所有需要的库
import pytest
import requests
import pandas as pd  # 用于读取Excel，替代xToolkit
import json  # 用于安全解析JSON格式的URL参数
import os  # 用于处理文件路径，生成报告路径


# 1. 读取 Excel 测试用例（处理格式转换和异常）
# 这个函数必须在它被调用的地方（所有测试用例 = read_test_cases("test.xls")）之前！
def read_test_cases(file_path):
    """
    从 Excel 读取测试用例，并处理 URL参数 的字符串转字典问题。
    增加错误处理，确保文件存在和格式正确。
    """
    try:
        # 读取 Excel (Pandas 兼容 .xls 和 .xlsx，需要 openpyxl 支持)
        df = pd.read_excel(file_path, sheet_name=0)  # sheet_name=0 表示读取第一个工作表
        cases = []
        for index, row in df.iterrows():
            case = row.to_dict()

            # 对从Excel读取到的值进行清理和类型转换
            # 确保关键字段不为 None/NaN，并去除前后空白
            case["接口URL"] = str(case.get("接口URL", "")).strip()
            case["请求方式"] = str(case.get("请求方式", "GET")).upper().strip()  # 默认GET，并转大写

            # 处理 URL参数：从字符串安全转换为字典
            # 考虑到Excel中可能为空白或非字符串类型
            params_str = case.get("URL参数", "{}")  # 默认给个空字典字符串，避免None
            if pd.isna(params_str) or not isinstance(params_str, str):
                case["URL参数"] = {}
            else:
                try:
                    # 使用 json.loads 进行安全解析，如果Excel里写的是 {'key':'value'} 这种单引号，
                    # 尝试替换成双引号来兼容，但最佳实践是Excel里直接写标准JSON格式的 "{}"
                    case["URL参数"] = json.loads(params_str.replace("'", '"'))
                except json.JSONDecodeError:
                    print(
                        f"警告: 第 {index + 2} 行的 URL参数 '{params_str}' 格式错误，已置为空字典。请确保是标准JSON格式（使用双引号）。")
                    case["URL参数"] = {}  # 转换失败则置空

            # 再次检查关键字段，如果URL为空，则跳过此用例
            if not case["接口URL"]:
                print(f"警告: 第 {index + 2} 行的 接口URL 为空，已跳过此用例。")
                continue

            cases.append(case)
        return cases
    except FileNotFoundError:
        print(f"错误: 文件 '{file_path}' 未找到。请确保 Excel 测试用例文件和脚本在同目录。")
        return []
    except Exception as e:
        # 捕获其他所有读取Excel可能出现的错误，提供明确提示
        print(f"Excel 文件读取失败，请检查文件内容、格式或权限: {e}")
        return []


# 2. 调用 read_test_cases 函数来加载测试用例
# 这行代码必须在 read_test_cases 函数定义之后！
所有测试用例 = read_test_cases("test.xls")

# 3. 核心逻辑：根据测试用例的加载情况，决定是否执行 pytest
if not 所有测试用例:
    print("没有找到任何有效的测试用例，无法执行接口测试。请检查 Excel 文件 'test.xls' 内容和路径。")
    import sys

    sys.exit(1)  # 以非零状态码退出，表示脚本执行有异常（没有测试可运行）


# 4. 定义参数化测试函数
# 这个函数必须在 pytest.main() 调用之前被定义！
@pytest.mark.parametrize("case", 所有测试用例)
def test_api_request(case):
    """
    发送 HTTP 请求并进行基础断言。
    每个 Excel 行被参数化为一个 'case'。
    """
    print(f"\n--- 正在执行用例: {case.get('接口URL', 'N/A')} ---")

    # 从当前测试用例中提取参数
    url = case.get("接口URL")
    method = case.get("请求方式", "GET").upper()  # 确保是大写
    params = case.get("URL参数", {})  # 默认为空字典

    # 发送 HTTP 请求并处理异常
    try:
        response = requests.request(
            method=method,
            url=url,
            params=params,  # 对于 GET/URL 参数，使用 params
            timeout=15  # 设置超时时间，防止长时间阻塞
        )
        response.raise_for_status()  # 如果状态码是 4xx 或 5xx，会抛出 HTTPError

        print(f"响应状态码: {response.status_code}")
        # 打印响应内容片段，方便调试
        print(f"响应内容片段: {response.text[:200]}...")  # 只打印前200字符

    except requests.exceptions.Timeout:
        pytest.fail(f"接口请求超时: {url}")
    except requests.exceptions.ConnectionError:
        pytest.fail(f"接口连接失败，请检查网络或URL是否可达: {url}")
    except requests.exceptions.HTTPError as e:
        # 更详细地展示HTTP错误
        pytest.fail(f"接口返回HTTP错误 {e.response.status_code}: {url}, 响应体: {e.response.text[:200]}")
    except Exception as e:
        pytest.fail(f"接口请求发生未知错误: {str(e)}")

    # 基础断言：检查 HTTP 状态码是否为 200
    assert response.status_code == 200, f"状态码异常，预期 200，实际 {response.status_code}"
    # 这里可以根据需要添加更多断言，例如：
    # assert "success" in response.text, "响应内容中未包含 'success' 关键词"
    # if response.json(): # 假设响应是JSON
    #     assert response.json().get("code") == 0, "响应JSON的code字段不是0"


# 5. 脚本的入口点：当作为主程序运行时，执行 pytest
if __name__ == '__main__':
    # 定义测试报告的文件名
    report_file_name = "test_report.html"
    # 获取当前脚本文件所在的目录的绝对路径
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接出完整的报告文件路径
    report_full_path = os.path.join(current_script_dir, report_file_name)

    print(f"测试报告将尝试生成在: {report_full_path}")

    # 运行 pytest 命令，包含生成 HTML 报告的参数
    # -v: 详细模式
    # --color=yes: 启用彩色输出
    # --html={report_full_path}: 指定 HTML 报告的输出路径和文件名
    # --self-contained-html: 将所有样式和脚本嵌入到单个 HTML 文件中，方便分享
    try:
        pytest.main(["-v", "--color=yes", f"--html={report_full_path}", "--self-contained-html"])
        print(f"\n测试执行完毕。请检查生成的 '{report_file_name}' 文件。")
    except Exception as e:
        print(f"\n执行 pytest 过程中发生错误: {e}")