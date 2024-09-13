import os
import time
import random
import datetime
import urllib.parse
import pandas as pd
from threading import Thread, BoundedSemaphore, Lock
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import tkinter as tk
import logging

# 全局变量和锁
progress_dict = {}
progress_lock = Lock()
exit_flag = False
excel_path = "C:\\Users\\yuesen\\Desktop\\爬虫参数.xlsx"


def update_progress(course, page, total_pages, crawled):
    with progress_lock:
        progress_dict[course] = (page, total_pages, crawled)


# def is_course_crawled(course):
#     try:
#         df = pd.read_excel(excel_path)
#         crawled_status = df.loc[df['course'] == course, 'crawled'].values[0]
#         return crawled_status == 1
#     except Exception as e:
#         print(f"检查crawled状态时发生错误：{e}")
#         return False


def print_progress(text_widget):
    try:
        while True:
            time.sleep(1)  # 每1秒打印一次进度
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with progress_lock:
                text_widget.config(state=tk.NORMAL)
                text_widget.delete(1.0, tk.END)
                text_widget.insert(tk.END, f"------------{current_time}---------------------------\n")
                sorted_courses = sorted(progress_dict.keys())
                for course in sorted_courses:
                    page, total_pages, crawled = progress_dict[course]
                    percentage = (page / total_pages) * 100 if total_pages else 0

                    # 如果crawled为1，强制设为100%
                    if crawled == 1:
                        percentage = 100

                    if int(percentage) == 100 or page == total_pages or crawled == 1:
                        text_widget.insert(tk.END, f"线程: {course}, 爬取进度: {page}/{total_pages}， {percentage:.2f}%\n", "green")
                    elif 0 < percentage < 100:
                        text_widget.insert(tk.END, f"线程: {course}, 爬取进度: {page}/{total_pages}， {percentage:.2f}%\n", "yellow")
                    else:
                        text_widget.insert(tk.END, f"线程: {course}, 爬取进度: {page}/{total_pages}， {percentage:.2f}%\n", "red")
                    logging.info(f"线程{course}: {page}/{total_pages}.")
                text_widget.insert(tk.END, "----------------------------------------------------------\n")
                text_widget.config(state=tk.DISABLED)
    except RuntimeError as e:
        logging.error(f"Error printing progress: {e}")
        print(f"线程异常退出: {e}")
    except Exception as e:
        logging.error(f"Error printing progress: {e}")
        print(f"线程异常退出: {e}")


def setup_text_widget_tags(text_widget):
    text_widget.tag_config("green", foreground="#006400", font=("Helvetica", 20))  # 字体大小可在此调整
    text_widget.tag_config("yellow", foreground="#CDBE70", font=("Helvetica", 20))  # 字体大小可在此调整
    text_widget.tag_config("red", foreground="#CD5C5C", font=("Helvetica", 20))  # 字体大小可在此调整


def update_crawled_status(course, output_widget):
    try:
        print(f"三秒后更新excel表格，请勿操作...")
        output_widget.insert(tk.END, f"三秒后更新excel表格，请勿操作...\n")
        output_widget.see(tk.END)
        time.sleep(3)
        df = pd.read_excel(excel_path)
        df.loc[df['course'] == course, 'crawled'] = 1
        df.to_excel(excel_path, index=False)
        # 同时更新 progress_dict 中的 crawled 状态
        with progress_lock:
            if course in progress_dict:
                page, total_pages, _ = progress_dict[course]
                progress_dict[course] = (page, total_pages, 1)
        logging.info(f"完成 {course}的爬取，更新Excel表格。")
        print(f"Excel表格已更新。")
        output_widget.insert(tk.END, f"Excel表格已更新。\n")
        output_widget.see(tk.END)
    except Exception as e:
        # 文件不存在或损坏时的处理
        logging.error(f"更新Excel表格时发生错误：{e}")
        if os.path.exists(excel_path):
            # 如果文件存在但无法访问，尝试删除并重新创建
            try:
                os.remove(excel_path)
                logging.info(f"Excel文件损坏，已删除：{excel_path}")
                # 可以在这里重新创建一个空的Excel文件
                pd.DataFrame().to_excel(excel_path, index=False)
            except Exception as delete_error:
                logging.error(f"删除Excel文件时发生错误：{delete_error}")
        else:
            logging.info(f"Excel文件不存在，将创建新的文件。")


def crawl_page(start_url, path, pages, file_name, course, semaphore, output_widget):
    with semaphore:
        if not os.path.exists(path):
            try:
                os.makedirs(path)
            except OSError as e:
                output_widget.insert(tk.END, f"Error: {e}. Directory '{path}' could not be created.\n")
                output_widget.see(tk.END)

        file_path = os.path.join(path, f"{file_name}.xlsx")
        # count = 1
        # while os.path.exists(file_path):
        #     file_path = os.path.join(path, f"{file_name}_{count}.xlsx")
        #     count += 1

        options = Options()
        options.headless = True
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--disable-background-networking")
        options.add_argument("--disable-sync")
        options.add_argument("--disable-translate")
        options.add_argument("--metrics-recording-only")

        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_argument("--log-level=3")
        # 上面三行实现了控制台不乱输出
        options.add_argument("--safebrowsing-disable-auto-update")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-software-rasterizer")
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--headless')

        all_links = []
        urls_to_visit = [start_url]

        driver = webdriver.Edge(options=options)
        try:
            for i in range(pages):
                current_url = urls_to_visit[i]
                driver.get(current_url)
                wait = WebDriverWait(driver, 15)
                max_retries = 4
                retries = 0

                while retries < max_retries:
                    try:
                        # print("正在等待元素加载以爬取链接...")
                        random_wait = random.uniform(0.5, 1)
                        stable_element = wait.until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="mic_cont_icon"]/div[2]/div')))
                        time.sleep(random_wait)
                        # print("元素加载完成。")
                        links = driver.find_elements(By.XPATH, '//*[@id="b_results"]/li/h2/a')

                        if len(links) == 0:
                            # print("未找到链接元素，刷新页面并重试...")
                            driver.refresh()
                            retries += 1
                            continue
                        else:
                            page_links = [link.get_attribute('href') for link in links if
                                          link.get_attribute('href') is not None]
                            all_links.extend(page_links)
                            # print(f"成功爬取第 {i + 1} 页的链接，数量为 {len(links)} 个。")
                            break

                    except TimeoutException:
                        output_widget.insert(tk.END, f"线程{course}等待超时，正在刷新页面重试...\n")
                        logging.warning(f"线程{course}超时，正在重试")
                        output_widget.see(tk.END)
                        retries += 1
                        driver.refresh()
                        time.sleep(random.uniform(0.5, 1))  # 随机等待时间，避免立即重试
                    except Exception as e:
                        output_widget.insert(tk.END, f"线程{course}遇到错误：{e}\n")
                        logging.warning(f"线程{course}遇到错误：{e}")

                if retries == max_retries:
                    output_widget.insert(tk.END, f"重试 {retries} 次后仍未找到链接元素，跳过当前页面。\n")
                    output_widget.insert(tk.END, f"线程{course}已无页面可爬取，程序退出。\n", "red")
                    output_widget.see(tk.END)
                    update_crawled_status(course, output_widget)
                    update_progress(course, i + 1, pages, 1)
                    break

                # 在更新进度的地方传递 crawled 参数
                update_progress(course, i + 1, pages, 0)

                try:
                    if not os.path.exists(file_path):
                        df = pd.DataFrame(all_links, columns=['URL'])
                        df.to_excel(file_path, index=False)
                    else:
                        existing_df = pd.read_excel(file_path)
                        new_df = pd.DataFrame(all_links, columns=['URL'])
                        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                        combined_df.to_excel(file_path, index=False)

                    output_widget.insert(tk.END, f"线程{course}成功保存{len(all_links)}个数据\n")
                    output_widget.see(tk.END)
                except Exception as e:
                    output_widget.insert(tk.END, f"线程{course}保存数据时发生错误：{e}\n")
                    logging.warning(f"线程{course}遇到错误：{e}")
                    output_widget.see(tk.END)

                if i < pages - 1:
                    parsed_url = urllib.parse.urlparse(current_url)
                    query_params = urllib.parse.parse_qs(parsed_url.query)
                    current_first_value = int(query_params.get('first', [0])[0])
                    if len(links) == 0:
                        next_first_value = current_first_value + 1
                    else:
                        next_first_value = current_first_value + len(links)

                    query_params['first'] = [str(next_first_value)]
                    new_url_query = urllib.parse.urlencode(query_params, doseq=True)
                    next_url = urllib.parse.urljoin(f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}",
                                                    '?' + new_url_query)
                    urls_to_visit.append(next_url)

                random_wait = random.uniform(0.1, 1)
                # print(f"等待{random_wait:.2f}秒后关闭浏览器，防止被识别为爬虫...")
                time.sleep(random_wait)
                all_links.clear()

                update_progress(course, i + 1, pages, 0)

            driver.quit()
            output_widget.insert(tk.END, f"线程{course}已完成。\n")
            output_widget.see(tk.END)
        except Exception as e:
            output_widget.insert(tk.END, f"线程 {course} 遇到错误：{e}\n")
            logging.warning(f"线程{course}遇到错误：{e}")
            output_widget.see(tk.END)
        finally:
            if exit_flag:
                output_widget.insert(tk.END, f"线程 {course} 接收到退出信号，将退出。\n")
                output_widget.see(tk.END)
            driver.quit()


def course_and_country(start_dict):
    course_dict = {}
    for key, value in start_dict.items():
        country = value['country']
        if country not in course_dict:
            course_dict[country] = []
        course_dict[country].append(key)
    return course_dict


def ed_course_and_country(ing_dict):
    ed_course_dict = {}
    for key, value in ing_dict.items():
        if value.get('crawled') == 1:
            continue
        else:
            country = value['country']
            if country not in ed_course_dict:
                ed_course_dict[country] = []
            ed_course_dict[country].append(key)
    return ed_course_dict


def read_parameters_from_excel(excel_path):
    if not os.path.exists(excel_path):
        raise FileNotFoundError("请指定有效的Excel文件路径。")
    df = pd.read_excel(excel_path)
    params = {}
    for index, row in df.iterrows():
        course = row['course']
        if not course:
            print(f"错误：'country'字段不能为空。跳过行 {index + 1}。")
            continue
        if "bing" not in row['start_url'] or "first" not in row['start_url']:
            print(f"错误：URL中必须包含 'bing' 和 'first' 参数。跳过行 {index + 1}。")
            continue
        if not row['file_name']:
            print(f"错误：请指定保存数据的文件名。跳过行 {index + 1}。")
            continue
        if not row['save_path'] or not os.path.isdir(row['save_path']):
            print(f"错误：无效的保存文件夹路径 '{row['save_path']}'")
            os.makedirs(row['save_path'], exist_ok=True)
            continue
        params[course] = {
            'crawled': row['crawled'],
            'country': row['country'],
            'start_url': row['start_url'],
            'pages': int(row['pages']),
            'file_name': row['file_name'],
            'path': row['save_path']
        }
    return params


def threaded_crawler(params, max_threads, output_widget):
    semaphore = BoundedSemaphore(max_threads)
    threads = []
    for course, details in params.items():
        if details['crawled'] == 0:
            logging.info(f"Starting thread for course {course}.")
            update_progress(course, 0, details['pages'], 0)
            t = Thread(target=crawl_page, args=(
                details['start_url'], details['path'], details['pages'], details['file_name'], course, semaphore, output_widget))
            threads.append(t)
            t.start()

    try:
        for thread in threads:
            thread.join()
    except KeyboardInterrupt:
        output_widget.insert(tk.END, "程序接收到退出信号，正在安全退出...\n")
        logging.info("Received keyboard interrupt, exiting safely.")
        output_widget.see(tk.END)
        global exit_flag
        exit_flag = True
        # 等待所有线程响应退出信号并安全退出
        for thread in threads:
            thread.join()
    finally:
        if exit_flag:
            output_widget.insert(tk.END, "所有线程已安全退出。\n")
            output_widget.see(tk.END)
            logging.info("All threads have exited safely.")


def main():
    max_threads = 6  # 设置最大线程数量
    params = read_parameters_from_excel(excel_path)

    # 创建GUI
    root = tk.Tk()
    root.title("多线程爬虫进度监控")

    # 第一部分：爬取进度
    frame1 = tk.Frame(root)
    frame1.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

    progress_text = tk.Text(frame1, width=80, height=30, state=tk.DISABLED)  # 取消滚动条，使用Text而不是ScrolledText
    progress_text.pack(fill=tk.BOTH, expand=True)

    # 设置文本框标签
    setup_text_widget_tags(progress_text)

    # 第二部分：其他输出
    frame2 = tk.Frame(root)
    frame2.pack(fill=tk.BOTH, expand=True, side=tk.RIGHT)

    output_text = tk.Text(frame2, width=40, height=30)  # 取消滚动条，使用Text而不是ScrolledText
    output_text.pack(fill=tk.BOTH, expand=True)

    # 启动进度打印线程
    progress_thread = Thread(target=print_progress, args=(progress_text,), daemon=True)
    progress_thread.start()

    # 启动爬虫线程
    crawler_thread = Thread(target=threaded_crawler, args=(params, max_threads, output_text))
    crawler_thread.start()

    root.mainloop()


if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(filename='crawler_2.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    main()
    print("所有进程均已退出。")
