import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging

WAIT_TIME = 20  # 延长等待时间到20秒，确保元素加载
logger = logging.getLogger("BookingChecker")  # 这行是核心，必须加！


def select_earliest_in_range(driver, start_date, end_date, daily_start_time, daily_end_time):
    """
    在指定日期和时间范围内选择最早可用的时段
    :param driver: WebDriver实例
    :param start_date: 开始日期（date对象）
    :param end_date: 结束日期（date对象）
    :param daily_start_time: 每天开始时间（time对象）
    :param daily_end_time: 每天结束时间（time对象）
    :return: 是否成功选择
    """
    try:
        # 获取所有行数据并筛选符合条件的时段
        rows = WebDriverWait(driver, WAIT_TIME).until(
            EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="slotSelectionForm:slotTable_data"]/tr'))
        )
        valid_slots = []

        for row in rows:
            # 提取时间文本（确保元素可见）
            time_label = WebDriverWait(row, WAIT_TIME).until(
                EC.visibility_of_element_located((By.XPATH, './td[2]/label'))
            )
            time_text = time_label.text.strip()

            # 解析时间
            try:
                slot_time = datetime.strptime(time_text, "%A, %d %B %Y %I:%M %p")
                slot_date = slot_time.date()
                slot_time_of_day = slot_time.time()

                # 检查是否在日期范围内
                if not (start_date <= slot_date <= end_date):
                    continue

                # 检查是否在每天的时间范围内
                if not (daily_start_time <= slot_time_of_day <= daily_end_time):
                    continue

                # 符合条件的时段加入列表
                valid_slots.append((slot_time, time_text, row))
            except Exception as e:
                logger.error(f"时间解析失败（{time_text}）：{e}")
                continue

        if not valid_slots:
            logger.info("指定范围内无符合条件的时段")
            return False, None

        # 按时间排序，选择最早的时段
        valid_slots.sort(key=lambda x: x[0])  # 按时间排序
        earliest_time, earliest_text, earliest_row = valid_slots[0]
        logger.info(f"\n最早可用时段：{earliest_text}")

        # 获取行属性
        data_ri = earliest_row.get_attribute("data-ri")
        data_rk = earliest_row.get_attribute("data-rk")

        # 定位目标行并激活
        row_xpath = f'//tbody[@id="slotSelectionForm:slotTable_data"]/tr[@data-ri="{data_ri}"]'
        target_row = WebDriverWait(driver, WAIT_TIME).until(
            EC.element_to_be_clickable((By.XPATH, row_xpath))
        )
        driver.execute_script("arguments[0].click();", target_row)
        time.sleep(0.5)

        # 修改隐藏字段值，模拟选中状态
        hidden_input = WebDriverWait(driver, WAIT_TIME).until(
            EC.presence_of_element_located((By.ID, "slotSelectionForm:slotTable_selection"))
        )

        driver.execute_script(f"arguments[0].value = '{data_rk}';", hidden_input)
        # 触发表单变更事件
        driver.execute_script("""
            var event = new Event('change', {bubbles: true});
            arguments[0].dispatchEvent(event);
        """, hidden_input)

        logger.info(f"已选择最早时段：{earliest_text}")
        print(f"已选择最早时段：{earliest_text}")
        return True, earliest_text  # 返回是否成功和选中的时间

    except Exception as e:
        logger.error(f"时间选择错误：{str(e)}")
        return False


def final_select_near_time(driver, target_time, tolerance):
    # 保留原函数用于兼容（如果需要）
    pass