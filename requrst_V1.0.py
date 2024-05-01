import csv
import requests
from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import pandas as pd
import os
import webbrowser
from urllib.parse import urlparse, urlunparse

paper = input("想收集的考试名称：")
classes = int(input("班级呢？："))
first = int(input("从第几个人开始"))
last = int(input("第几位结束？"))
answer1 = input ("是否清空？yes/no")
if answer1.lower() in ["y","yes"]:
    with open('student_scores.txt', 'w') as f:
        pass
    with open('student_scores.csv', 'w') as f:
        pass
print("正在调取信息")
print("正在收集 ",classes," 班级从第 ",first,"到",last," 同学的 ",paper," 分数明细表 ")
start_time = time.time()
for i in range (first,last+1):
    formatted_i = '%02d' % i
    # 初始化Edge浏览器驱动
    driver = webdriver.Edge()

    # 打开网页
    driver.get("http://58.213.73.26:37068/Home/login")

    # from selenium.webdriver.common.by import By

    # 假设 driver 已经是一个正确初始化的 WebDriver 实例
    username_input = driver.find_element(By.CSS_SELECTOR, "input.form-control[placeholder='请输入用户名']")
    username_input.send_keys("lz",classes,formatted_i)

    # 等待密码输入框加载
    password_input = driver.find_element(By.CSS_SELECTOR, "input.form-control[placeholder='输入密码']")
    password_input.send_keys("lz123456!")

    # 等待登录按钮加载
    login_button = WebDriverWait(driver, 8).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@type='button']"))
    )
    login_button.click()

    # 假设你已经定位到了链接元素
    # 等待元素可见
    try:
        link = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "学情分析系统"))
        )
        # 点击链接
        link.click()
        # 如果点击后需要等待页面加载或执行某些动作，可以在这里添加相应的代码
        print("点击了链接学情分析系统")
    except Exception as e:
        print(f"无法定位或点击链接：{e}")

    try:
        # 定位链接元素
        link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, paper))
        )
        # 点击链接
        link.click()
        print("点击了链接 ",paper)
    except Exception as e:
        print(f"无法定位或点击链接：{e}")
    # 爬取网站内容
    url = driver.current_url
    page_source = driver.page_source

    # 使用BeautifulSoup解析页面源代码
    soup = BeautifulSoup(page_source, 'html.parser')
    table = soup.find('table', {'class': 'tab_report stuReport', 'id': 'tab_grcjd'})

    # 找到表头和表格行
    headers = [th.get_text(strip=True) for th in table.find('thead').find_all('th')]
    rows = table.find('tbody').find_all('tr')

    # 提取每个学生的成绩信息
    student_scores = []
    for row in rows:
        # 获取表格中的每个单元格数据
        cells = row.find_all('td')
        # 提取并整理学生的姓名和各科成绩
        scores = {header: cell.get_text(strip=True) for header, cell in zip(headers, cells)}
        student_scores.append(scores)

        # 打印每个学生的成绩信息
    for score in student_scores:
        print(f"姓名: {score['姓名']}")
        print(f"考号: {score['考号']}")
        print(f"总分: {score['总分']}")
        for subject, grade in score.items():
            if subject in ['语文', '数学', '英语', '物理', '化学', '生物', '政治', '地理', '历史']:
                print(f"{subject}: {grade}")
    print()

    csv_filename = 'student_scores.csv'
    with open(csv_filename, 'a', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.DictWriter(csvfile, fieldnames=headers)
        # 如果文件是空的，则写入表头
        if csvfile.tell() == 0:
            csvwriter.writeheader()
        for score in student_scores:
            csvwriter.writerow(score)

        # 追加到文本文件
    txt_filename = 'student_scores.txt'
    with open(txt_filename, 'a', encoding='utf-8') as txtfile:
        for score in student_scores:
            line = ', '.join(f'{k}: {v}' for k, v in score.items())
            # 在每条成绩记录前添加换行符以确保新记录在新的一行
            txtfile.write(str(i))
            txtfile.write(line + '\n')
    print(f"成绩已保存到 {csv_filename} 和 {txt_filename}")
    driver.quit()
else:
    print("没人了，一共记录了：",last,"人")

# 读取CSV文件
df = pd.read_csv('student_scores.csv')

# 替换null值为0，以便进行数学计算
df.replace('null', 0, inplace=True)

# 将总分列中的字符串转换为浮点数
df['总分'] = pd.to_numeric(df['总分'], errors='coerce')

# 计算所有学生的平均分（不包括总分列和考号列）
average_scores = df.drop(['姓名', '考号', '总分'], axis=1).mean()

# 计算所有学生的总分（仅包括语文到历史这些科目）
df['总分'] = df[['语文', '数学', '英语', '物理', '化学', '生物', '政治', '地理', '历史']].sum(axis=1)

# 计算所有学生的平均总分
average_total_score = df['总分'].mean()

# 分析哪个学生总分最高和最低
highest_score_student = df.loc[df['总分'].idxmax(), '姓名']
lowest_score_student = df.loc[df['总分'].idxmin(), '姓名']

# 分析哪个科目平均分最高和最低（不包括总分列）
highest_avg_subject = average_scores.idxmax()
lowest_avg_subject = average_scores.idxmin()

# 如果需要，还可以计算其他统计信息，比如标准差等
standard_deviations = df[['语文', '数学', '英语', '物理', '化学', '生物', '政治', '地理', '历史']].std()

# 输出统计信息到控制台
print("所有科目的平均分:")
print(average_scores)
print("\n所有学生的平均总分:")
print(average_total_score)
print("\n总分最高的学生是:", highest_score_student)
print("总分最低的学生是:", lowest_score_student)
print("\n平均分最高的科目是:", highest_avg_subject)
print("平均分最低的科目是:", lowest_avg_subject)
print("\n每科的标准差:")
print(standard_deviations)

# 将总结数据转换为DataFrame
summary_df = pd.DataFrame({
    '科目': average_scores.index,
    '平均分': average_scores.values
})

# 将总结数据追加到原始DataFrame的末尾（注意：这不是直接修改原CSV文件）
df_with_summary = pd.concat([df, summary_df]).reset_index(drop=True)

# 将带有总结的DataFrame保存为新的CSV文件
df_with_summary.to_csv('student_scores_with_summary.csv', index=False)

# 将总结信息追加到文本文件中
summary_text = f"所有科目的平均分如下：\n"
for index, score in average_scores.items():
    summary_text += f"{index}: {score:.2f}\n"
summary_text += f"\n所有学生的平均总分是: {average_total_score:.2f}\n"
summary_text += f"总分最高的学生是: {highest_score_student}\n"
summary_text += f"总分最低的学生是: {lowest_score_student}\n"
summary_text += f"平均分最高的科目是: {highest_avg_subject}\n"
summary_text += f"平均分最低的科目是: {lowest_avg_subject}\n"

'''print("所有科目的平均分:")
print(average_scores)
print("\n所有学生的平均总分:")
print(average_total_score)
print("\n总分最高的学生是:", highest_score_student)
print("总分最低的学生是:", lowest_score_student)
print("\n平均分最高的科目是:", highest_avg_subject)
print("平均分最低的科目是:", lowest_avg_subject)
print("\n每科的标准差:")
print(standard_deviations)'''

# 将总结文本追加到指定的文本文件末尾
with open('summary.txt', 'a') as f:
    f.write(summary_text)

end_time = time.time()
execution_time = end_time - start_time
script_dir = os.path.dirname(os.path.abspath(__file__))
# 将文件路径转换为file://格式的URL
parsed = urlparse('file:' + os.path.abspath(script_dir))
file_url = urlunparse(parsed)



print("已经爬取全部内容")
print(f"一共执行了 {execution_time:.6f} 秒.")
time.sleep(3)
print("正在为你打开文件夹...")
# 使用webbrowser打开目录
webbrowser.open(file_url)