# -*- coding: utf-8 -*-
import base64
import os
import sys
import time

import tkinter as tk
import tkinter.messagebox as msg
import tkinter.ttk as ttk

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from subprocess import CREATE_NO_WINDOW

from webdriver_manager.chrome import ChromeDriverManager

VERSION = "订餐小助手V1.3"

date = time.strftime("%Y-%m-%d ", time.localtime())


class DinnerHelper:
    def __init__(self) -> None:
        self.root = tk.Tk()

        self.name = tk.StringVar(value="")
        self.email = tk.StringVar(value="")
        self.department = tk.StringVar(value="")
        self.password = tk.StringVar(value="")
        self.want_sp_set = tk.BooleanVar(value=False)
        self.is_debug_mode = tk.BooleanVar(value=False)

        self.status = tk.StringVar(value="")
        self.progress = tk.IntVar(value=0)

        self.config = os.path.expanduser("~/.dinner_helper.config")
        self.is_ordered = False
        self.is_open_site = False
        self.is_found_config = False
        self.need_reopen = False

        if os.path.exists(self.config):
            self.is_found_config = True
            self.update_status("找到配置文件")
            with open(self.config, "rb") as f:
                s = f.readline()
                s = base64.b64decode(s).decode()
                s = s.split('$')
                self.name.set(s[0])
                self.email.set(s[1])
                self.department.set(s[2])
                self.password.set(s[3])
        else:
            self.update_status("初次运行未设置个人信息")

    def open_browser(self):
        options = webdriver.ChromeOptions()
        service = Service(ChromeDriverManager().install())
        if not self.is_debug_mode.get():
            options.add_argument('headless')
            service.creationflags = CREATE_NO_WINDOW

        self.driver = webdriver.Chrome(service=service, chrome_options=options)

        self.driver.get(
            "https://ovtcn.sharepoint.cn/teams/Admin/Lists/List3/AllItems.aspx")

    def open_site(self):
        self.update_status("打开浏览器", 0)
        self.open_browser()

        self.update_status("输入邮箱", 35)
        WebDriverWait(self.driver, 10).until(
            ec.presence_of_element_located((By.ID, "i0116")))
        self.driver.find_element(By.ID, "i0116").send_keys(self.email.get())
        self.driver.find_element(By.ID, "i0116").send_keys(Keys.ENTER)

        self.update_status("输入密码", 40)
        WebDriverWait(self.driver, 10).until(
            ec.presence_of_element_located((By.ID, "passwordInput")))
        self.driver.find_element(
            By.ID, "passwordInput").send_keys(self.password.get())
        self.driver.find_element(
            By.ID, "passwordInput").send_keys(Keys.ENTER)

        try:
            self.driver.find_element(By.ID, "idSIButton9").click()
        except Exception:
            self.need_reopen = True
            msg.showwarning("错误", "无法登录，请检查密码是否正确")

        self.update_status("登录成功", 50)

    def order(self):
        self.open_site()

        self.update_status("新建条目", 60)
        self.driver.find_element(By.ID, "idHomePageNewItem").click()

        self.update_status("输入名字", 70)
        self.driver.find_element(
            By.ID, "Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField").send_keys(self.name.get())

        self.update_status("输入部门", 75)
        self.driver.find_element(
            By.ID, "_x90e8__x95e8__eb2a7400-e297-48f2-a420-6dc7b5526048_$TextField").send_keys(self.department.get())

        self.update_status("是否选择特色餐", 80)
        if self.want_sp_set.get():
            self.driver.find_element(
                By.ID, "_x5957__x9910__x7c7b__x522b__x00_645f5e81-e442-4042-904d-448484c70478_$RadioButtonChoiceField1").click()

        self.update_status("输入公司", 85)
        dropdown = self.driver.find_element(
            By.ID, "_x516c__x53f8__f03bcd26-e326-40a0-b55e-633bac828fb2_$DropDownChoice")
        dropdown.find_element(By.XPATH, "//option[. = 'OV']").click()

        self.update_status("输入工作地点", 90)
        dropdown = self.driver.find_element(
            By.ID, "_x529e__x516c__x5730__x70b9__a18fb9ad-1b6b-49c1-8c09-edd334ff9d60_$DropDownChoice")
        dropdown.find_element(By.XPATH, "//option[. = '上海张江']").click()

        self.update_status("订餐", 95)
        self.driver.find_element(
            By.ID, "ctl00_ctl33_g_a0c338ef_96f8_45ea_aee5_4819043d0ca8_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()
        self.driver.quit()

        self.is_ordered = True
        self.update_status(p=100)
        msg.showinfo("一键订餐", time.strftime("%Y-%m-%d %H:%M:%S ") + "订餐成功")

    def unorder(self):
        self.open_site()

        self.update_status("展开公司", 60)
        self.driver.find_element(By.CSS_SELECTOR, ".ms-gb > a").click()
        self.update_status("展开套餐", 70)
        self.driver.find_element(By.CSS_SELECTOR, ".ms-gb2 > a").click()
        self.update_status("选择订餐", 80)

        self.driver.find_element(By.CSS_SELECTOR, ".ms-core-overlay").click()
        self.driver.find_element(By.CSS_SELECTOR, ".ms-ellipsis-icon").click()
        self.update_status("删除订餐", 90)
        self.driver.find_element(By.ID, "ID_DeleteItem").click()
        self.driver.quit()

        self.is_ordered = False
        self.update_status(p=100)
        msg.showinfo("取消订餐", time.strftime("%Y-%m-%d %H:%M:%S ") + "取消成功")

    def LineUp(self):
        self.driver.get(
            "https://ovtcn.sharepoint.cn/teams/Admin/Lists/List3/view.aspx")

        a = self.driver.find_element(By.CSS_SELECTOR, "#titl1-1_2_ > tr > td > span")
        pass


    def gui(self):
        self.root.title("订餐小助手")
        self.root.iconbitmap(self.get_path("./food-tray.ico"))
        self.root.geometry(self.get_position(300, 170))
        self.root.resizable(False, False)

        menubar = tk.Menu(self.root)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="个人信息", command=self.settings)
        settings_menu.add_checkbutton(
            label="调试模式", variable=self.is_debug_mode)

        settings_menu.add_separator()
        settings_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="设置", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)

        help_menu.add_command(
            label="文档", command=lambda: msg.showerror("抱歉", "还未实现的功能"))
        help_menu.add_command(label="关于", command=lambda: msg.showinfo(
            title="关于", message=VERSION + "\nrj.xu@ovt.com"))
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

        ttk.Label(self.root, text="今天需要订晚餐吗？").pack()
        ttk.Button(self.root, text="取消订餐",
                   command=self.unorder).pack()
        ttk.Button(self.root, text="一键订餐", command=self.order).pack()

        ttk.Checkbutton(self.root, text="是否选择特色餐？",
                        variable=self.want_sp_set).pack()

        ttk.Separator(self.root).pack()
        ttk.Progressbar(self.root, variable=self.progress,
                        length=250, mode='determinate').pack()

        tk.Label(self.root, textvariable=self.status,
                 bd=1, relief=tk.SUNKEN, anchor=tk.E).pack(side=tk.BOTTOM, fill=tk.X)

        if not self.is_found_config:
            msg.showwarning("提示", "初次运行请先填写设置-个人信息")
        self.root.mainloop()

    def settings(self):
        self.top = tk.Toplevel()
        self.top.title("设置")
        self.top.iconbitmap(self.get_path("./food-tray.ico"))
        self.top.geometry(self.get_position(200, 120))
        self.top.resizable(False, False)

        frame = tk.Frame(self.top)

        ttk.Label(frame, text="中文名").grid(row=0, column=0)
        ttk.Entry(frame, textvariable=self.name).grid(row=0, column=1)

        ttk.Label(frame, text="邮箱").grid(row=1, column=0)
        ttk.Entry(frame, textvariable=self.email).grid(row=1, column=1)

        ttk.Label(frame, text="部门").grid(row=2, column=0)
        ttk.Entry(frame, textvariable=self.department).grid(row=2, column=1)

        ttk.Label(frame, text="密码").grid(row=3, column=0)
        ttk.Entry(frame, textvariable=self.password,
                  show="*").grid(row=3, column=1)
        frame.pack()

        ttk.Separator(self.top).pack()

        ttk.Button(self.top, text="保存", command=lambda: (
            self.save_settings(), self.update_status("更新账户设置", 0), self.top.destroy())).pack()

    def get_position(self, w, h):
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        return "%dx%d+%d+%d" % (w, h, x, y)

    def get_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def update_status(self, s=None, p=0):
        if s:
            self.status.set(date + s)
        else:
            self.status.set(date + ("订餐成功" if self.is_ordered else "未订餐"))
        self.progress.set(p)
        self.root.update()

    def save_settings(self):
        s = self.name.get() + '$' + self.email.get() + '$' + \
            self.department.get() + '$' + self.password.get()
        s_base64 = base64.b64encode(s.encode())
        with open(self.config, "wb") as f:
            f.write(s_base64)


if __name__ == "__main__":
    d = DinnerHelper()
    d.gui()
