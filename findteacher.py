from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import re

def convert_names_from_excel(filepath):
    # 载入Excel文件，确保第一行数据不被用作列名
    df = pd.read_excel(filepath, header=None)
    # 移除空行
    df.dropna(inplace=True)
    # 定义一个函数来提取姓名
    def extract_name(text):
        match = re.search(r'\[(.*?)\]', text)
        return match.group(1) if match else None
    # 应用函数，提取姓名
    names = df.iloc[:, 0].apply(extract_name).dropna().tolist()
    return names

def search_teacher_info(name, driver):
    # 打开指定的URL
    driver.get("https://faculty.xidian.edu.cn/jscx.jsp?urltype=tree.TreeTempUrl&wbtreeid=1005")

    # 设置显式等待
    wait = WebDriverWait(driver, 3)  # 最多等待3秒

    # 找到姓名输入框并输入搜索名字，等待元素可见
    name_input = wait.until(EC.visibility_of_element_located((By.ID, "teacherNameu8")))
    name_input.clear()
    name_input.send_keys(name)

    # 点击查询按钮，等待元素可点击
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'search-btn') and contains(text(), '查询')]")))
    search_button.click()

    # 等待搜索结果加载，这里等待的是结果区域的某个标志性元素出现
    wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@class='text-title']")))

    # 分析搜索结果，获取教师信息
    results = driver.find_elements(By.XPATH, "//div[@class='text-title']")
    for result in results:
        teacher_name = result.find_element(By.XPATH, "./div[@class='name']/span/a").text
        department = result.find_element(By.XPATH, "./div[@class='dw']").text.split("：")[-1]
        if teacher_name == name:
            return department
    return "未找到"

def main():
    # 使用转换函数读取并处理名单数据
    names = convert_names_from_excel("名单表.xlsx")

    # 启动Chrome浏览器
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # 启用无头模式
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        results = {}
        for name in names:
            department = search_teacher_info(name, driver)
            if department not in results:
                results[department] = [name]
            else:
                results[department].append(name)

        # 转换结果为DataFrame
        max_len = max(len(v) for v in results.values())
        for key in results.keys():
            results[key] += [""] * (max_len - len(results[key]))

        df_results = pd.DataFrame(results)
        # 保存到新的Excel表格
        df_results.to_excel("搜索结果.xlsx", index=False)
    finally:
        driver.quit()

# 运行主函数
if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    print("程序运行时间：", end_time - start_time, "s")
