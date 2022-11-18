import os
import time
import pickle
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
try:
    import init_env
except:
    pass


class SeleniumEmail:
    def __init__(self, exe_path=None, hidden=False):
        options = Options()
        # 隐藏窗口
        if hidden:
            options.add_argument("--headless")
        exe_path = exe_path or 'C://chromedriver.exe'
        self.wd = webdriver.Chrome(executable_path=exe_path, options=options)
        self.wd.implicitly_wait(10)
        self.login_url = 'https://mail.163.com'

    # 载入cookies
    def _load_cookies(self):
        cookies = pickle.load(open('cookies.pkl', 'rb'))
        for cookie in cookies:
            _dict = {
                'domain': 'mail.163.com',
                'name': cookie.get('name'),
                'value': cookie.get('value')
            }
            self.wd.add_cookie(_dict)
        self.wd.refresh()

    # 手动登录
    def _login_hand(self, username, password):
        self.wd.switch_to.frame(self.wd.find_element_by_tag_name('iframe'))
        # 自动填充邮箱
        element = self.wd.find_element_by_xpath('//input[@name="email"]')
        element.clear()
        element.send_keys(username)
        time.sleep(0.5)

        # 自动填充密码
        element = self.wd.find_element_by_xpath('//input[@name="password"]')
        element.clear()
        element.send_keys(password)
        time.sleep(1)

        # 手动登录 防止验证码
        # 等待用户点击登录，100秒超时
        print('请手动点击登录按钮')
        for i in range(100):
            time.sleep(1)
            current_url = str(self.wd.current_url).rstrip('/')
            if current_url != self.login_url:
                time.sleep(1)
                pickle.dump(self.wd.get_cookies(), open('cookies.pkl', 'wb'))
                break

    # 登录
    def login(self, username, password):
        self.wd.get(self.login_url)
        if os.path.exists('cookies.pkl'):
            self._load_cookies()
        current_url = str(self.wd.current_url).rstrip('/')
        if current_url == self.login_url:
            print('cookie过期，手动登录')
            self._login_hand(username, password)
        current_url = str(self.wd.current_url).rstrip('/')
        if current_url == self.login_url:
            print('登录失败，退出程序')
            exit(0)
        print('登录成功')

    # 查看最新邮件
    def look_newest_email(self):
        # 点击收件箱
        recv_ems = self.wd.find_elements_by_class_name('nui-tree-item-text')
        recv_ems[0].click()
        time.sleep(1)

        # 点击最新邮件
        em = self.wd.find_elements_by_class_name('nui-txt-flag0')
        em[0].click()

    # 下载附件(浏览器默认下载方式 不能指定位置)
    def download_files_bak(self):
        ems = self.wd.find_elements_by_class_name('lh0')
        time.sleep(1)
        for i, em in enumerate(ems):
            # 显示出下载框
            self.wd.execute_script("arguments[0].setAttribute(arguments[1],arguments[2])", em, 'class',
                                   'lh0 nui-bdr-item nui-bg-item gh0 fM0')

            time.sleep(1)
            # 点击下载
            files = self.wd.find_element_by_class_name('qs0.nui-fClear').find_elements_by_xpath(
                '//a[@target="downloadFrame"]')
            files[i + 1].click()
            time.sleep(1)

            # 隐藏下载框
            self.wd.execute_script("arguments[0].setAttribute(arguments[1],arguments[2])", em, 'class', 'lh0')
            time.sleep(1)

    # 下载附件(可指定保存位置 推荐)
    def download_email_files(self):
        files = self.wd.find_element_by_class_name('qs0.nui-fClear').find_elements_by_xpath(
            '//a[@target="downloadFrame"]')
        # 获取附件文件名列表
        filenames = [item.text for item in self.wd.find_elements_by_class_name('dh0')]
        cookies = self.wd.get_cookies()
        req_cookie = {}
        for ck in cookies:
            req_cookie[ck.get('name')] = ck.get('value')
        for i, file in enumerate(files):
            if i == 0:
                continue
            url = file.get_attribute('href')
            print('download url=', url)
            res = requests.get(url, cookies=req_cookie)
            path = f"download/{filenames[i - 1]}"
            with open(path, 'wb') as f:
                f.write(res.content)

    # 发送邮件
    def send_email(self, recv_email, subject_name, content, files=None):
        # 点击写信
        em = self.wd.find_element_by_class_name('js-component-component.ra0.mD0')
        em.click()
        # 输入收件人
        print('填充收件人')
        self.wd.find_element_by_class_name('nui-editableAddr-ipt').send_keys(recv_email)
        # 输入主题
        print('填充主题')
        ems = self.wd.find_elements_by_class_name('nui-ipt-input')
        for em in ems:
            id = em.get_attribute('id')
            if 'subjectInput' in id:
                em.send_keys(subject_name)
        time.sleep(1)
        # 输入内容
        print('填充内容')
        self.wd.switch_to.frame(self.wd.find_element_by_class_name('APP-editor-iframe'))
        self.wd.find_element_by_tag_name('p').send_keys(content)
        # 切回主iframe
        self.wd.switch_to.parent_frame()
        # 上传附件
        print('上传附件')
        files = files or []
        for file_path in files:
            self.wd.find_element_by_class_name('O0').send_keys(file_path)
            time.sleep(0.5)
        time.sleep(1)
        # 点击发送邮件
        print('点击发送邮件')
        ems = self.wd.find_elements_by_class_name('nui-btn-text')
        for em in ems:
            if em.text == '发送':
                em.click()
                break


if __name__ == '__main__':
    # 设置测试邮箱账号密码
    email_username = 'xxxxxx@163.com'
    email_password = 'xxxxxx'

    # 通过环境变量方式设置账号密码（可忽略）
    if os.environ.get('email_username') and os.environ.get('email_password'):
        email_username = os.environ.get('email_username')
        email_password = os.environ.get('email_password')

    # 设置测试邮件发送信息
    recv_email = '1186156343@qq.com'
    subject_name = '测试主题'
    content = '我是内容哈哈哈'

    # 示例上传文件
    examples = os.listdir('examples')
    base_path = os.path.join(os.path.dirname(__file__), 'examples')
    files = [os.path.join(base_path, item) for item in examples]

    se = SeleniumEmail()

    print('开始登录')
    se.login(email_username, email_password)

    print('查看最新邮件')
    se.look_newest_email()

    # print('下载附件')
    # se.download_email_files()

    print('发送测试邮件')
    se.send_email(recv_email, subject_name, content, files)
    print('发送完成')

