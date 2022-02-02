# -*- coding: utf-8 -*-
"""爬取微信公众号的文章"""

from selenium import webdriver
import time
import requests
import re
import random
import os
import pdfkit  # pdf提取存储工具库

# 设置headers
header = {
    "HOST": "mp.weixin.qq.com",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; OWW64; rv:53.0) Gecko/20100101 Firefox/53.0"
}

# 当前代码文件所在文件夹地址
cur_path = os.path.dirname(__file__)

# 设置要爬取的公众号列表
lt_account_names = ['量子位', 'DeepTech深科技']


class WechatMP(webdriver.Firefox):
    """Firefox的子类"""

    def __init__(self):
        """
        初始化微信公众号编辑平台页面
        :return: webdriver.Firefox
        """
        # 设置浏览器启动参数
        browser_options = webdriver.FirefoxOptions()
        # browser_options.add_argument('--headless')  # 隐藏窗体（不能隐藏，要扫描微信公众平台二维码登录！）

        # 继承父类的构造方法
        super(WechatMP, self).__init__(options=browser_options)

        # 设置相关参数
        self.implicitly_wait(5)  # 等待元素出现的时间最大为5s

        # 定义属性
        self.__dict_cookies = None
        self.__token = None
        self.is_login = False

    def login(self):
        """登录微信公众平台"""
        # 进行扫码登录
        input('请扫码登录任意公众号账户！完成后回到控制台界面 ENTER 继续进程')
        self.is_login = True

        '''微信公众平台目前每次登录都需要进行扫码验证，所以在此处就不做账号和密码+cookies保存的无效操作了'''
        print()  # 控制台展示目的的空行

    def start(self):
        """启动微信公众号平台"""
        # 微信公众平台的链接
        url = 'https://mp.weixin.qq.com/'
        # 打开链接
        self.get(url=url)

        # 如果尚未登录，则进行登录操作；反之，则刷新至主页即可
        if not self.is_login:
            self.login()

    def get_cookies_dict(self):
        """
        获取字典(dict)形式的cookies
        :return:dict
        """
        """"""
        # 如果已有dict_cookies值，则直接返回结果即可
        if self.__dict_cookies is not None:
            return self.__dict_cookies

        # 声明一个dict 存储结果
        dict_cookies = {}
        # 获取列表形式的cookies
        cookie_items = self.get_cookies()

        # 获取到的cookies是列表形式，将cookies转成json形式
        for cookie_item in cookie_items:
            dict_cookies[cookie_item['name']] = cookie_item['value']

        # 设置为当前对象的属性
        self.__dict_cookies = dict_cookies

        # 返回dict结果
        return dict_cookies

    def get_token(self):
        """
        获取登录后的token信息
        实现思路：登录成功后，输入微信公众号网页（https://mp.weixin.qq.com/），页面链接为自动更新为带token的形式（e.g. https://mp.weixin.qq.com/cgi-bin/home?t=home/index&lang=zh_CN&token=51146553），通过正则表达式提取其中的token信息
        :return:str
        """
        # 如果已有token值，则直接返回结果即可
        if self.__token is not None:
            return self.__token

        # 刷新至公众号主页
        self.start()  # 启动微信公众平台页面
        time.sleep(2)  # 等待2s的页面加载

        # 通过正则表达式获取token信息
        # 当前url样例：https://mp.weixin.qq.com/cgi-bin/home?t=home/index&lang=zh_CN&token=51146553
        cur_url = self.current_url  # 当前页面的链接地址
        token_id = re.findall(r'token=(\d+)', str(cur_url))[0]

        # 设置为当前对象的属性
        self.__token = token_id

        return token_id


class Article:
    """公众号文章类"""

    def __init__(self, source_account, aid, copyright_type, cover_link, create_time, digest, link, title, update_time):
        """
        公众号文章属性初始化
        :param source_account:str(未来可能用Account对象，但目前还没有需求）
        :param aid:str
        :param copyright_type:int(0:原创, 1:原创, 2:转载)
        :param cover_link:str
        :param create_time:int
        :param digest:str 短句概括
        :param link:str
        :param title:str
        :param update_time:int
        """
        self.source_account = source_account
        self.aid = aid
        self.copyright_type = copyright_type
        self.cover_link = cover_link
        self.create_time = create_time
        self.digest = digest
        self.link = link
        self.title = title
        self.update_time = update_time

    def save_to_disc_as_pdf(self):
        """以pdf的形式存储至本地"""
        # 创建保存文件夹
        # 1、先创建母目录“公众号文章爬取合集”
        if not os.path.exists(os.path.join(cur_path, '公众号文章爬取合集')):
            os.mkdir(os.path.join(cur_path, '公众号文章爬取合集'))

        # 2、再创建以“{公众号}”为名的子目录
        if not os.path.exists(os.path.join(cur_path, '公众号文章爬取合集', self.source_account)):
            os.mkdir(os.path.join(cur_path, '公众号文章爬取合集', self.source_account))

        # 构建输出路径的字符串
        output_path = '%s/%s.pdf' % (os.path.join(cur_path, '公众号文章爬取合集', self.source_account), str(self))  # 输出的文件绝对路径

        try:
            # 配置wkhtmltopdf.exe位置
            config = pdfkit.configuration(wkhtmltopdf=r'%s\工具合集\wkhtmltox\bin\wkhtmltopdf.exe' % cur_path)
            # 输入html链接，输出：pdf文件
            pdfkit.from_string(input=self.link, output_path=output_path, configuration=config)
            # 控制台提示
            print('%s By %s，输出成功！' % (self.title, self.source_account))
        except OSError:
            print(
                '文章保存失败！失败原因：你没有安装wkhtmltopdf工具，需要安装该工具（安装地址：https://github.com/JazzCore/python-pdfkit/wiki/Installing-wkhtmltopdf），并配置bin文件夹路径为环境变量')
            # 原文提示
            print(
                'If this file exists please check that this process can read it or you can pass path to it manually in method call, check README. Otherwise please install wkhtmltopdf - https://github.com/JazzCore/python-pdfkit/wiki/Installing-wkhtmltopdf')
            print()  # 控制台展示目的

            # 传递OSError错误
            raise OSError

    def __str__(self):
        """公众号的str输入格式"""
        # 格式化[创建时间]和[更新时间]
        create_time_f = time.strftime('%Y年%m月%d日', time.localtime(self.create_time))
        update_time_f = time.strftime('%Y年%m月%d日', time.localtime(self.update_time))

        return '{title}：{link}（{account}于{create_time}发布，最近一次更新时间：{update_time}）'.format(title=self.title,
                                                                                         link=self.link,
                                                                                         account=self.source_account,
                                                                                         create_time=create_time_f,
                                                                                         update_time=update_time_f)


class Account:
    """公众号类"""

    def __init__(self, name):
        """
        公众号初始化方法
        :param name:公众号名称
        """
        self.name = name
        self.__fake_id = None  # 微信公众号平台搜索中的公众号id
        self.__contents = None  # 微信公众号平台的文章内容
        self.__cnt_articles = None  # 微信公众号的文章总数

    def get_fake_id(self, wechat_mp=None):
        """
        获取公众号的fakeid
        :param wechat_mp:WechatMP (default:None)
        :return:str
        """
        # 如果已有fake_id值，则直接返回结果即可
        if self.__fake_id is not None:
            return self.__fake_id

        # 通过公众号平台搜索获取
        '''1、登入公众号搜索接口'''
        # 搜索微信公众号的接口地址
        search_url = 'https://mp.weixin.qq.com/cgi-bin/searchbiz?'
        # 获取token信息
        token = wechat_mp.get_token()
        # 构建微信公众号搜索接口输入的参数：有三个变量 -> 微信公众号token、随机数random、搜索的微信公众号名字
        query_id = {
            'action': 'search_biz',
            'token': token,  # 登录携带的token信息
            'lang': 'zh_CN',
            'f': 'json',
            'ajax': '1',
            'random': random.random(),  # 随机数
            'query': self.name,  # 进行搜索的微信公众号名称
            'begin': '0',
            'count': '5'
        }

        # 获取dict形式的cookies
        dict_cookies = wechat_mp.get_cookies_dict()
        # 打开搜索微信公众号接口地址，需要传入相关参数信息如：cookies、params、headers
        # response的样例：{'base_resp': {'ret': 0, 'err_msg': 'ok'}, 'list': [{'fakeid': 'MzIzNjc1NzUzMw==', 'nickname': '量子位', 'alias': 'QbitAI', 'round_head_img': 'http://mmbiz.qpic.cn/mmbiz_png/YicUhk5aAGtCEFSVW5ubo08Zfv1qB5iapricibTBdETkBNtolJxnSUib6UXhjWWz3aib8vETY00P2lKR1uG3qLHicSoWg/0?wx_fmt=png', 'service_type': 1, 'signature': '追踪人工智能新趋势，关注科技行业新突破'}, {'fakeid': 'MzkyMzIzMjU0OA==', 'nickname': '量子位点', 'alias': '', 'round_head_img': 'http://mmbiz.qpic.cn/mmbiz_png/Sicic1OQOtC111vJbaG1UBiakzlTBHLAlW7VK6v51mvbN4ruWh2SwUe2BPceg4PUKznGJ57jANU5UpzBCpEjMa35A/0?wx_fmt=png', 'service_type': 1, 'signature': '一个对物理感兴趣的人，不定期写文章科普物理'}, {'fakeid': 'MzUyNTcxMzk3MA==', 'nickname': '浙江孚天量子科技有限公司', 'alias': 'zjft2018', 'round_head_img': 'http://mmbiz.qpic.cn/mmbiz_png/tXrEWpHHlcn2h15tLu8ibMmOLicRdZYMhWXHNqYwcSgsQNyvd7jFgpccDdhd1JMMXl2HniaJe8MIXsN1OzhDibt7icA/0?wx_fmt=png', 'service_type': 2, 'signature': '本公司以健康养生为理念，是一家集研发、生产、销售为一体的高科技现代企业。其产品通过国家有关权威检测认证，拥有技术应用研究、技术成果转换、数据植入核心科技。'}, {'fakeid': 'MzU1NzU5MjY3Ng==', 'nickname': '量子时空商城', 'alias': 'lzsk66666666', 'round_head_img': 'http://mmbiz.qpic.cn/mmbiz_png/ewDOa9wuLOpYmw6cibzmF7aBJd8mSekCsMaSM3Zoibibgyiaxvib4L18kYaVsVaicfV549bD208qN2brxY5A0QCV8Rgw/0?wx_fmt=png', 'service_type': 1, 'signature': '此公众号使命是： 1、传达李终辉先生一手信息、意见，做到李先生和大家真正同心同行。 2、量子文化道场，弘扬量子文化。 3、量子人成长基地，帮量子人茁壮成长。 4、后期会有广告位供供应商投放。'}], 'total': 4}
        search_response = requests.get(search_url, cookies=dict_cookies, headers=header, params=query_id)

        '''2、搜索目标公众号'''
        print('\n正在进行以关键词【%s】的公众号搜索...' % self.name)
        # 搜索结果列表
        lt_search_results = search_response.json().get('list')
        # 取搜索结果中的第一个公众号
        # 结果样例：{'fakeid': 'MzIzNjc1NzUzMw==', 'nickname': '量子位', 'alias': 'QbitAI', 'round_head_img': 'http://mmbiz.qpic.cn/mmbiz_png/YicUhk5aAGtCEFSVW5ubo08Zfv1qB5iapricibTBdETkBNtolJxnSUib6UXhjWWz3aib8vETY00P2lKR1uG3qLHicSoWg/0?wx_fmt=png', 'service_type': 1, 'signature': '追踪人工智能新趋势，关注科技行业新突破'}
        first_result = lt_search_results[0]
        # 提取该结果的fakeid
        fake_id = first_result['fakeid']
        print('%s的fakeid为：%s' % (first_result['nickname'], fake_id))

        # 将搜索获得的fakeid储存至当前对象
        self.__fake_id = fake_id

        return fake_id

    def get_cnt_articles(self, wechat_mp=None, is_update=False):
        """
        获取当前公众号的文章总数
        :param wechat_mp:WechatMP
        :param is_update:bool 是否更新公众号的文章总数
        :return:int
        """
        # 如果已经获取过文章总数数据且不是更新操作，则直接返回结果即可
        if self.__cnt_articles is not None and not is_update:
            return self.__cnt_articles

        # 获取通用参数
        token = wechat_mp.get_token()  # 当前登录的token
        dict_cookies = wechat_mp.get_cookies_dict()  # 当前登录状态的dict形式cookies
        fake_id = self.get_fake_id(wechat_mp=wechat_mp)  # 目标公众号的标识码

        # 微信公众号文章接口地址
        app_msg_url = 'https://mp.weixin.qq.com/cgi-bin/appmsg?'
        # 搜索文章需要传入几个参数：登录的公众号token、要爬取文章的公众号fakeid、随机数random
        query_id_data = {
            'token': token,
            'lang': 'zh_CN',
            'f': 'json',
            'ajax': '1',
            'random': random.random(),
            'action': 'list_ex',
            'begin': '0',  # 不同页，此参数变化，变化规则为每页加5
            'count': '5',
            'query': '',
            'fakeid': fake_id,
            'type': '9'
        }
        # 打开搜索的微信公众号文章列表页
        # response样例：{'app_msg_cnt': 2668, 'app_msg_list': [{'aid': '2649651589_1', 'album_id': '0', 'appmsg_album_infos': [], 'appmsgid': 2649651589, 'checking': 0, 'copyright_type': 1, 'cover': 'https://mmbiz.qlogo.cn/mmbiz_jpg/JJtKEey0hPaiarMPxIMla0NQej0coKfa7pOZAbsNdB1zkyFKIJLcribyrobdM4XFUd7QicY9ib73biaZuvJmIJSL5yQ/0?wx_fmt=jpeg', 'create_time': 1643687857, 'digest': '“做世界最准的钟”。', 'has_red_packet_cover': 0, 'is_pay_subscribe': 0, 'item_show_type': 0, 'itemidx': 1, 'link': 'http://mp.weixin.qq.com/s?__biz=MzA3NTIyODUzNA==&mid=2649651589&idx=1&sn=e09fef4e50351f90072bd4faa201f3f6&chksm=87696d9cb01ee48a762a248cb28a6e4774612900d79aca2028ee7f628eb9fd63b3bf5e6308d5#rd', 'media_duration': '0:00', 'mediaapi_publish_status': 0, 'pay_album_info': {'appmsg_album_infos': []}, 'tagid': [], 'title': '专访光学巨擘叶军：造出世界上最精确的原子光钟！可帮助做出更好的GPS，也可用于重新测量宇宙常数', 'update_time': 1643687856},], 'base_resp': {'err_msg': 'ok', 'ret': 0}}
        app_msg_response = requests.get(app_msg_url, cookies=dict_cookies, headers=header, params=query_id_data)

        # 获取文章总数
        cnt_articles = app_msg_response.json()['app_msg_cnt']
        # 存储该结果为__cnt_articles
        self.__cnt_articles = cnt_articles

        # 控制台提示结果
        print('%s有%i篇文章' % (self.name, cnt_articles))

        return cnt_articles

    def get_content(self, wechat_mp=None, is_update=False):
        """
        获取单个公众号的文章内容
        :param wechat_mp:WechatMP
        :param is_update:bool 是否更新公众号的文章存储内容
        :return:list
        """
        # 如果已有存储的文章内容且不是更新操作，则直接返回结果即可
        if self.__contents is not None and not is_update:
            return self.__contents

        # 微信公众号文章接口地址
        app_msg_url = 'https://mp.weixin.qq.com/cgi-bin/appmsg?'

        # 获取通用参数
        token = wechat_mp.get_token()  # 当前登录的token
        dict_cookies = wechat_mp.get_cookies_dict()  # 当前登录状态的dict形式cookies
        fake_id = self.get_fake_id(wechat_mp=wechat_mp)  # 目标公众号的标识码
        cnt_articles = self.get_cnt_articles(wechat_mp=wechat_mp)  # 目标公众号的文章总数

        # 请求获取该公众号的文章
        query_id_data = {
            'token': token,
            'lang': 'zh_CN',
            'f': 'json',
            'ajax': '1',
            'random': random.random(),
            'action': 'list_ex',
            'begin': 0,  # begin的存储顺序为先进后出——“最近一篇为0”
            'count': '10',
            'query': '',
            'fakeid': fake_id,
            'type': '9'
        }
        # 发起请求，app_msg_response存储请求回复的结果
        # response样例：{'app_msg_cnt': 2668, 'app_msg_list': [{'aid': '2649651589_1', 'album_id': '0', 'appmsg_album_infos': [], 'appmsgid': 2649651589, 'checking': 0, 'copyright_type': 1, 'cover': 'https://mmbiz.qlogo.cn/mmbiz_jpg/JJtKEey0hPaiarMPxIMla0NQej0coKfa7pOZAbsNdB1zkyFKIJLcribyrobdM4XFUd7QicY9ib73biaZuvJmIJSL5yQ/0?wx_fmt=jpeg', 'create_time': 1643687857, 'digest': '“做世界最准的钟”。', 'has_red_packet_cover': 0, 'is_pay_subscribe': 0, 'item_show_type': 0, 'itemidx': 1, 'link': 'http://mp.weixin.qq.com/s?__biz=MzA3NTIyODUzNA==&mid=2649651589&idx=1&sn=e09fef4e50351f90072bd4faa201f3f6&chksm=87696d9cb01ee48a762a248cb28a6e4774612900d79aca2028ee7f628eb9fd63b3bf5e6308d5#rd', 'media_duration': '0:00', 'mediaapi_publish_status': 0, 'pay_album_info': {'appmsg_album_infos': []}, 'tagid': [], 'title': '专访光学巨擘叶军：造出世界上最精确的原子光钟！可帮助做出更好的GPS，也可用于重新测量宇宙常数', 'update_time': 1643687856},], 'base_resp': {'err_msg': 'ok', 'ret': 0}}
        app_msg_response = requests.get(app_msg_url, cookies=dict_cookies, headers=header, params=query_id_data)
        # 解析response中返回的文章列表
        lt_articles = app_msg_response.json()['app_msg_list']

        # 声明文章对象的存储列表
        lt_contents = []
        # 遍历返回的所有文章，并将其存储为Article对象，执行...操作
        for json_article in lt_articles:
            # 根据json信息构建Article对象
            article = Article(source_account=self.name, aid=json_article['aid'],
                              copyright_type=json_article['copyright_type'], cover_link=json_article['cover'],
                              create_time=json_article['create_time'], digest=json_article['digest'],
                              link=json_article['link'], title=json_article['title'],
                              update_time=json_article['update_time'])

            # 存储文章
            try:
                # 临时方案：以pdf形式保存至本地
                article.save_to_disc_as_pdf()

            except OSError:
                print('文章保存为pdf失败...正在退出操作...\n\n')
                return None

            # 将Article对象储存至该公众号对象的内容列表中
            lt_contents.append(article)

        # 赋值当前对象的内容列表
        self.__contents = lt_contents
        # 展示目的：结尾在控制台中打印空行
        print('%s成功读取%i篇文章。（已保存）\n\n' % (self.name, len(lt_contents)))

        return lt_contents


if __name__ == '__main__':
    # 初始化浏览器
    obj_wechat_mp = WechatMP()
    # 登录微信公众号平台
    obj_wechat_mp.start()

    # 构建公众号为对象，并储存至字典(key:name, value:obj)
    dict_accounts = {}

    # 遍历要爬取的公众号列表
    for account_name in lt_account_names:
        # 控制台提示当前的公众号操作
        print(('收集公众号【%s】的文章' % account_name).center(50, '-'))

        # 建立公众号对象
        obj_account = Account(name=account_name)

        # 获取该公众号的文章
        contents = obj_account.get_content(wechat_mp=obj_wechat_mp)

        # 将对象存入字典中
        dict_accounts[account_name] = obj_account
