import os
import shutil
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException

# 设置下载路径
download_path = r"D:\移动终端广分互联网\下载"


# 打开对应网站并导出数据
def download_file(driver, file_name, path):
    try:
        # 查找并点击下载按钮
        try:
            # 跳转到当前导出文件的页面
            driver.get("https://jsls.jinritemai.com/mfa/organization-management/" + path)

            # 等待并点击导出按钮
            # search_button = driver.find_element(By.XPATH, '//*[@id="_yylx_organization-management"]'
            #                                                 '/div/div/div/div/div[2]/div/div/div[1]/button')
            export_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,
                            '//*[@id="_yylx_organization-management"]/div/div/div/div/div[3]/div[2]/div[2]/button')))
            export_button.click()
            print("点击导出按钮")
            time.sleep(3)

            # 点击查看记录按钮
            records_button = driver.find_element(By.XPATH, '//*[@id="_yylx_organization-management"]'
                                                           '/div/div/div/div/div[3]/div[2]/div[3]/span/button')
            records_button.click()
            time.sleep(3)

        except TimeoutException:
            print("没有找到元素，判断为子账号管理界面")
            # 判断已授权or邀请中
            if file_name == "导出子账号数据_邀请中":
                invite_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,
                            '//*[@id="inner_tab_status-tab-1"]')))
                invite_button.click()
                time.sleep(3)
            export_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,
                    '//*[@id="rc-tabs-0-panel-sub-account-manage"]/div/div[2]/div[2]/div[1]/div/div[3]/div[1]/button/span')))
            export_button.click()
            print("点击导出按钮")
            time.sleep(3)

            # 点击查看记录按钮
            records_button = driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-sub-account-manage"]'
                                                        '/div/div[2]/div[2]/div[1]/div/div[3]/div[2]/span/button/span')
            records_button.click()
            time.sleep(3)

        except Exception as e:
            # 捕获其他非超时异常并报错
            print(f"发生错误: {e}")
            driver.quit()
            raise  # 重新抛出异常以结束进程

        # 等待刚刚导出的数据准备好下载按钮并点击
        download_button = WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH,
                            '//*[@id="root"]/div/div/div[2]/div/div/div/div/div/table/tbody/tr[2]/td[7]/div/div/a')))
        download_button.click()

        # 等待文件下载完成
        time.sleep(10)

        # 获取文件并改名移动覆盖前数据源数据
        file_conversion(file_name)

    except Exception as e:
        print(type(e))
        raise  # 重新抛出异常以结束进程


def file_conversion(file_name):
    # 获取下载目录中的最新文件名
    files = os.listdir(download_path)
    files = [f for f in files if os.path.isfile(os.path.join(download_path, f))]  # 只获取文件
    if files:
        latest_file = max([os.path.join(download_path, f) for f in files], key=os.path.getctime)  # 获取最新文件
        print(f"最新下载的文件是: {latest_file}")
    else:
        print("没有下载文件。")

    # 修改文件名称并覆盖原文件
    data_source = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\数据源"
    data_path = os.path.join(data_source, f"{file_name}.xlsx")
    # 如果源文件已存在，先删除后移动
    if os.path.exists(data_path):
        os.remove(data_path)
    # 将导出的文件重命名
    renamed_file = os.path.join(download_path, f"{file_name}.xlsx")
    os.rename(latest_file, renamed_file)
    # 移动重命名后的文件到数据源文件夹
    shutil.move(renamed_file, data_source)
    print(f"{renamed_file} 已成功移动并覆盖到 {data_source}")


def export_data():
    # 清空下载文件夹
    shutil.rmtree(download_path)  # 删除目录及其所有内容
    os.makedirs(download_path)  # 重新创建空目录
    # 配置 Chrome 下载路径
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_path,  # 设置下载目录
        "download.prompt_for_download": False,        # 禁用下载提示
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    # 启用无头模式
    # chrome_options.add_argument("--headless")  # 启用无头模式
    # 设置用户数据目录
    chrome_options.add_argument(r'--user-data-dir=C:\Users\83825\AppData\Local\Google\Chrome\User Data')
    # 排除自动化的标记
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])

    # 启动 Chrome 浏览器
    driver = webdriver.Chrome(options=chrome_options)

    # 打开目标网页
    driver.get("https://jsls.jinritemai.com/mfa/homepage")

    # 等待页面加载
    time.sleep(3)

    # 下载文件目录
    download_dist = {'导出管理单元数据': 'company/list', '导出门店数据': 'store/list',
                     '导出子账号数据_已授权': 'account', '导出子账号数据_邀请中': 'account'}

    # 循环导出四份数据
    for d in download_dist:
        download_file(driver, d, download_dist[d])
        print(d + "成功！")

    # 关闭浏览器
    driver.quit()


