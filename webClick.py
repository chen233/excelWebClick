import undetected_chromedriver as uc
from selenium.common import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import timeSelect
import logging
from selenium.webdriver.chrome.service import Service  # 新增：导入Service类
# 首先在文件顶部添加邮件相关依赖
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

# 在webClick.py中添加邮件配置（请根据实际情况修改）
EMAIL_CONFIG = {
    "smtp_server": "smtp.qq.com",  # 例如：smtp.qq.com、smtp.gmail.com
    "smtp_port": 587,  # 通常为587或465
    "smtp_username": "583134449@qq.com",
    "smtp_password": "hbognhcftpudbcad",  # 邮箱密码或授权码
    "sender": "583134449@qq.com",
    "receiver": "Hungryworld0001@gmail.com"  # 接收通知的邮箱
}


# 添加发送邮件函数
def send_booking_success_email(config_data, booking_time):
    """发送预约成功邮件"""
    try:
        # 创建邮件对象
        msg = MIMEMultipart()
        msg['From'] = Header(f"预约系统 <{EMAIL_CONFIG['sender']}>")
        msg['To'] = EMAIL_CONFIG['receiver']
        msg['Subject'] = Header("【预约成功通知】驾驶考试预约已完成", 'utf-8')

        # 邮件内容
        content = f"""
        <p>驾驶考试预约已成功！</p>
        <p>预约信息如下：</p>
        <p>驾照编号：{config_data['dlNumber']}</p>
        <p>联系人：{config_data['contactName']}</p>
        <p>联系电话：{config_data['contactPhone']}</p>
        <p>考试类型：{config_data['Test type']}</p>
        <p>考试地点：{config_data['Region']} - {config_data['Centre']}</p>
        <p>预约时间：{booking_time}</p>
        <p>请按时前往参加考试。</p>
        """
        msg.attach(MIMEText(content, 'html', 'utf-8'))

        # 发送邮件
        server = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
        server.starttls()  # 启用TLS加密
        server.login(EMAIL_CONFIG['smtp_username'], EMAIL_CONFIG['smtp_password'])
        server.sendmail(
            EMAIL_CONFIG['sender'],
            EMAIL_CONFIG['receiver'],
            msg.as_string()
        )
        server.quit()

        logger.info("预约成功邮件已发送")
        return True
    except Exception as e:
        logger.error(f"发送邮件失败：{str(e)}")
        return False


TARGET_URL = "https://www.service.transport.qld.gov.au/SBSExternal/application/CleanBookingDE.xhtml"
WAIT_TIME = 10
logger = logging.getLogger("BookingChecker")  # 这行是核心，必须加！

# 从Excel读取的配置信息（需要在调用时传入）
# 这些字段需要与Excel中的配置对应
INPUT_DATA = {
    "dlNumber": "",  # 驾照编号
    "contactName": "",  # 联系人姓名
    "contactPhone": "",  # 联系电话
    "Test type": "",  # 考试类型
    "Region": "",  # 地区
    "Centre": "",  # 考试中心
    "contactEmail": "",  # 联系邮箱
    "CardNumber": "",  # 卡号
    "ExpiryMonth": "",  # 有效期月份
    "ExpiryYear": "",  # 有效期年份
    "CVN": ""  # 安全码
}


def openweb(start_date, end_date, daily_start_time, daily_end_time, config_data):
    """
    打开网页并完成预约流程
    :param start_date: 预约开始日期（date对象）
    :param end_date: 预约结束日期（date对象）
    :param daily_start_time: 每天开始时间（time对象）
    :param daily_end_time: 每天结束时间（time对象）
    :param config_data: 从Excel读取的配置信息字典
    """
    # 更新全局配置数据
    global INPUT_DATA
    INPUT_DATA.update(config_data)

    # 配置无头模式
    chrome_options = uc.ChromeOptions()
    # chrome_options.add_argument("--headless=new")  # 关键：启用无头模式,取消注释就不显示浏览器画面
    # chrome_options.add_argument('--start-maximized')  # 取消注释！启用浏览器最大化（普通模式生效）
    chrome_options.add_argument('--blink-settings=imagesEnabled=false')  # 禁用图片加载
    chrome_options.add_argument('--disable-extensions')  # 禁用扩展
    chrome_options.add_argument('--disable-gpu')  # 禁用GPU加速（非必要）
    chrome_options.add_argument('--no-sandbox')  # 禁用沙箱（提升速度）
    chrome_options.add_argument('--disable-dev-shm-usage')  # 避免共享内存不足
    chrome_options.add_argument('--ignore-certificate-errors')  # 忽略证书错误（加速加载）

    driver_service = Service(executable_path=r"chrome\chromedriver.exe")  # 注意：路径用r开头避免转义

    # 启动后台浏览器（其他代码不变）
    driver = uc.Chrome(service=driver_service, options=chrome_options)

    driver.get(TARGET_URL)  # 后续操作正常执行，无界面显示
    logger.info("已启动隐藏自动化特征的浏览器，正在访问网站...")
    try:
        # 点击继续按钮


        continue_btn = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "ui-button"))
        )
        continue_btn.click()

        # 输入驾照编号
        dl_input = WebDriverWait(driver, WAIT_TIME).until(
            EC.presence_of_element_located((By.ID, "CleanBookingDEForm:dlNumber"))
        )
        dl_input.clear()
        dl_input.send_keys(INPUT_DATA["dlNumber"])
        logger.info("已输入驾照编号")

        # 输入联系人姓名
        contact_input = driver.find_element(By.ID, "CleanBookingDEForm:contactName")
        contact_input.clear()
        contact_input.send_keys(INPUT_DATA["contactName"])
        logger.info("已输入联系人姓名")

        # 输入手机号
        contactPhone = driver.find_element(By.ID, "CleanBookingDEForm:contactPhone")
        contactPhone.clear()
        contactPhone.send_keys(INPUT_DATA["contactPhone"])
        logger.info("已输入手机号")

        # 选择考试类型
        try:
            dropdown_button = WebDriverWait(driver, WAIT_TIME).until(
                EC.element_to_be_clickable((By.ID, "CleanBookingDEForm:productType"))
            )
            dropdown_button.click()

            target_option = WebDriverWait(driver, WAIT_TIME).until(
                EC.visibility_of_element_located((
                    By.XPATH,
                    f'//ul[@id="CleanBookingDEForm:productType_items"]/li[text()="{INPUT_DATA["Test type"]}"]'
                ))
            )

            actions = ActionChains(driver)
            actions.move_to_element(target_option).click().perform()
            logger.info("考试类型选择成功！")
        except Exception as e:
            logger.info(f"考试类型选择失败：{e}")
            raise

        # 继续到下一步
        dropdown_button1 = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-success"))
        )
        dropdown_button1.click()
        dropdown_button2 = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.ID, "BookingConfirmationForm:actionFieldList:confirmButtonField"
                                               ":confirmButton"))
        )
        dropdown_button2.click()
        # 选择地区
        try:
            dropdown_button = WebDriverWait(driver, WAIT_TIME).until(
                EC.element_to_be_clickable((By.ID, "BookingSearchForm:region_label"))
            )
            dropdown_button.click()
            target_option = WebDriverWait(driver, WAIT_TIME).until(
                EC.visibility_of_element_located((
                    By.XPATH,
                    f'//ul[@id="BookingSearchForm:region_items"]/li[text()="{INPUT_DATA["Region"]}"]'
                ))
            )

            actions = ActionChains(driver)
            actions.move_to_element(target_option).click().perform()
            logger.info("地区选择成功！")
        except Exception as e:
            logger.info(f"地区选择失败：{e}")
            raise

        # 选择考试中心
        try:
            time.sleep(1)
            dropdown_button = WebDriverWait(driver, WAIT_TIME).until(
                EC.element_to_be_clickable((By.ID, "BookingSearchForm:centre"))
            )
            dropdown_button.click()
            target_option = WebDriverWait(driver, WAIT_TIME).until(
                EC.visibility_of_element_located((
                    By.XPATH,
                    f'//ul[@id="BookingSearchForm:centre_items"]/li[text()="{INPUT_DATA["Centre"]}"]'
                ))
            )

            actions = ActionChains(driver)
            actions.move_to_element(target_option).click().perform()
            logger.info("考试中心选择成功！")
        except Exception as e:
            logger.info(f"考试中心选择失败：{e}")
            raise

        # 继续到时间选择页面
        continue_btn = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-success"))
        )
        continue_btn.click()

        # 等待预约表格加载
        logger.info("\n等待预约时间表格加载...")
        WebDriverWait(driver, WAIT_TIME).until(
            EC.visibility_of_element_located((By.ID, "slotSelectionForm:slotTable"))
        )
        # 调用时间选择函数，选择范围内最早的时间
        success, selected_time = timeSelect.select_earliest_in_range(
            driver=driver,
            start_date=start_date,
            end_date=end_date,
            daily_start_time=daily_start_time,
            daily_end_time=daily_end_time
        )
        if not success:
            logger.info(f"{datetime.now()} 未在指定范围内找到可用时段")
            print(f"{datetime.now()} 未在指定范围内找到可用时段")
            return False
        # 成功选中后进入下一页
        continue_btn5 = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.ID, "slotSelectionForm:actionFieldList:confirmButtonField:confirmButton"))
        )
        continue_btn5.click()

        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.ID, "BookingConfirmationForm:actionFieldList:confirmButtonField:confirmButton")))
        element.click()  # 此时元素一定存在于当前DOM中
        time.sleep(1)
        # 填写邮箱
        contactEmail = driver.find_element(By.ID,"paymentOptionSelectionForm:paymentOptions:emailAddressField:emailAddress")
        contactEmail.clear()
        contactEmail.send_keys(INPUT_DATA["contactEmail"])
        logger.info("已输入邮箱")
        continue_btn1 = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-success"))
        )
        continue_btn1.click()
        time.sleep(1)

        # 填写付款信息
        contactCard = driver.find_element(By.ID, "CardNumber")
        contactCard.clear()
        contactCard.send_keys(INPUT_DATA["CardNumber"])

        contactMonth = driver.find_element(By.ID, "ExpiryMonth")
        contactMonth.clear()
        contactMonth.send_keys(INPUT_DATA["ExpiryMonth"])

        contactYear = driver.find_element(By.ID, "ExpiryYear")
        contactYear.clear()
        contactYear.send_keys(INPUT_DATA["ExpiryYear"])

        contactCVN = driver.find_element(By.ID, "CVN")
        contactCVN.clear()
        contactCVN.send_keys(INPUT_DATA["CVN"])
        logger.info("已填写付款信息，准备提交付款")
        # ---------------------- 新增：付款按钮点击 + 结果判断 ----------------------
        payment_success = False  # 付款结果标识（默认失败）
        try:
            # 1. 点击付款审核按钮（btnReviewPayment）
            continue_btn = WebDriverWait(driver, WAIT_TIME).until(
                EC.element_to_be_clickable((By.ID, "btnReviewPayment"))  # 等待按钮可点击，避免未加载完成
            )
            continue_btn.click()

            logger.info("已点击付款审核按钮，等待付款结果...")
            continue_btn1 = WebDriverWait(driver, WAIT_TIME).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='PAY']"))  # 精准匹配“PAY”文本的按钮
            )
            continue_btn1.click()
            # 2. 判断是否付款成功（核心：等待成功标识元素，超时则视为失败）
            # 若能走到这一步，说明成功找到成功元素
            logger.info("付款成功！已完成预约流程")
            payment_success = True
            send_booking_success_email(config_data, earliest_text)  # 发送成功邮件

        except Exception as e:
            # 捕获超时/元素不存在异常，视为付款未成功
            logger.error(f"付款流程异常：{str(e)}")
            # 可额外判断是否存在“付款失败”提示（可选，进一步细化失败原因）
            try:
                # 假设失败提示元素为 id="paymentFailMsg"
                fail_msg = driver.find_element(By.ID, "paymentFailMsg").text
                logger.error(f"付款失败，网站提示：{fail_msg}")
            except:
                logger.error("未明确获取付款失败原因，视为付款未完成")
            payment_success = False

        finally:
            # 原有：最后一步按钮点击（通常是“确认”或“完成”，即使付款失败也可能需要点击关闭）
            try:
                # 等待最后一步按钮可点击（根据实际ID调整，若没有可删除）
                final_btn = WebDriverWait(driver, WAIT_TIME).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "button"))
                )
                final_btn.click()
                time.sleep(5)  # 等待页面最终处理
            except Exception as e:
                logger.warning(f"最后一步按钮点击异常：{str(e)}")

            # 关闭浏览器（原有逻辑）
            driver.quit()

        # ---------------------- 关键：返回付款结果（决定Excel状态） ----------------------
        # 返回值说明：
        # - True：付款成功 → Excel状态为“已完成付款”
        # - False：付款未成功 → Excel状态为“已进入付款页面但未付款成功”
        return payment_success



    except Exception as e:
        logger.info(f"脚本出错：{e}")
        return False
    finally:
        driver.quit()
