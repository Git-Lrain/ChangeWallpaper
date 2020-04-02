import urllib
import ssl
from bs4 import BeautifulSoup
import os
import shutil
from PIL import Image
import getpass
import time
# 导入证书
ssl._create_default_https_context = ssl._create_stdlib_context


# 获取图片格式
def PicFormat_Get(file):
    img = Image.open(file)
    return(img.size)


# 获取文件信息
def FileInfo_Get(path):
    mtime = os.path.getmtime(path)
    return(mtime)


# 获取图片绝对路径
def Path_Get(path):
    path = os.path.abspath(path)
    return(path)


# 获取系统用户名
def Username_Get():
    username = getpass.getuser()
    return(username)


# 获取当前时间
def Time_Get():
    Ctime = time.strftime("%Y-%m-%d", time.localtime())
    return(Ctime)


# 重启资源管理器
def Restart_Explorer():
    os.system('taskkill /f /im explorer.exe')
    os.system('start explorer.exe')
    return 0


# 获取网页源码
def webContent_Get():
    Bing_Url = 'https://bing.ioliu.cn/'  # Bing壁纸的url
    header = {'User-Agent': '*'}  # 设置robots的header
    with urllib.request.urlopen(urllib.request.Request(url=Bing_Url, headers=header)) as code:
        data = code.read()  # 获取网页源码
        return(data.decode())  # 将源码解码后返回


# 查找网页源码中的图片链接
def picUrl_Capture(content):
    soup = BeautifulSoup(content, features="lxml")
    url = soup.find('img').get('src')  # 查找到第一张图片的缩略图链接
    tmp = url.split('/')[-1].split('_')
    name = tmp[0] + '_' + tmp[1]
    url = 'https://bing.ioliu.cn/photo/' + name + '?force=download'  # 拼接图片下载链接
    return(url)


# 下载图片
def Pic_Download(url, path):
    header = {'User-Agent': '*'}  # 设置robots的header
    with urllib.request.urlopen(urllib.request.Request(url=url, headers=header)) as f:
        with open(path, 'wb') as p:
            p.write(f.read())
            return(path)


# 获取Windows聚焦壁纸存放路径
def FocusPath_Find(username):
    path_f = 'C:/Users/' + username + '/AppData/Local/Packages'
    folder_name = ''
    folder_list = os.listdir(path_f)
    for i in folder_list:  # 不同版本的系统的后缀可能会不同
        if 'ContentDeliveryManager' in i:
            folder_name = i
    path_s = path_f + '/' + folder_name + '/LocalState/Assets/'
    return(path_s)


# Focus最新壁纸查找
def Focus_Find(path):
    focusPC_list = {}  # 存放适用于1920*1080的图片信息
    focusPH_list = {}  # 存放适用于1080*1920的图片信息
    pic_list = os.listdir(path)  # 获取指定目录下的文件
    for p in pic_list:
        tp = path + p  # 图片的完整路径
        if PicFormat_Get(tp)[0] == 1920:  # 判断图片的宽度是否为 1920 px
            focusPC_list[p] = FileInfo_Get(tp)  # 存储图片名和图片修改时间到字典中
        else:
            focusPH_list[p] = FileInfo_Get(tp)
    latestPC_pic = sorted(focusPC_list.items(), key=lambda kv: (kv[1], kv[0]))[-1][0]  # 选出字典中最新的图片的图片名
    latestPH_pic = sorted(focusPH_list.items(), key=lambda kv: (kv[1], kv[0]))[-1][0]
    return(latestPC_pic, latestPH_pic)


# Focus文件转储
def Pic_save(sourcepath, filename, pcpic_path, kind=0):
    # pcpic_path = './FocusPc/'+filename+'.jpg'
    phpic_path = './FocusPhone/'+filename+'.jpg'
    if kind == 0:
        shutil.copyfile(sourcepath+filename, pcpic_path)  # 复制Windows聚焦中的壁纸到指定文件夹下
        return(os.path.abspath(pcpic_path))  # 返回复制后图片的地址
    elif kind == 1:
        shutil.copyfile(sourcepath+filename, phpic_path)
        return 0
        # return(os.path.abspath(phpic_path))


# 应用壁纸
def Doing(filePath, username):
    Default_path = 'C:/Users/' + username + \
        '/AppData/Roaming/Microsoft/Windows/Themes/TranscodedWallpaper'
    Default_path_Cache = 'C:/Users/' + username + \
        '/AppData/Roaming/Microsoft/Windows/Themes/CachedFiles/CachedImage_1920_1080_POS4.jpg'
    shutil.copyfile(filePath, Default_path)  # 将准备要作为壁纸的文件复制到windows壁纸目录下
    shutil.copyfile(filePath, Default_path_Cache)


# 执行
def Bing_Doing():
    if not os.path.exists('Bing'):  #
        os.mkdir('Bing')
    filename = Time_Get()
    path = './Bing/' + filename + '.jpg'  # 将下载的文件以日期来命名
    if not os.path.exists(path):
        """
        可防止一日内多次启动电脑导致频繁刷新
        但若是一日内从Bing切换至Focus在切换至Bing
        则第二次切换至Bing时壁纸将无法及时应用
        因为此时文件夹中已有当日的壁纸文件
        """
        content = webContent_Get()
        url = picUrl_Capture(content)
        Pic_Download(url, path)
        abspath = Path_Get(path)
        username = Username_Get()
        Doing(abspath, username)
        Restart_Explorer()


# Windows壁纸应用
def Focus_Doing():
    if not os.path.exists('FocusPc'):
        os.mkdir('FocusPc')
    if not os.path.exists('FocusPhone'):
        os.mkdir('FocusPhone')
    username = Username_Get()
    path = FocusPath_Find(username)
    pic1, pic2 = Focus_Find(path)
    pcpic_path = './FocusPc/' + pic1 + '.jpg'
    if not os.path.exists(pcpic_path):
        apath = Pic_save(path, pic1, pcpic_path, 0)
        Pic_save(path, pic2, pcpic_path, 1)
        if FileInfo_Get(pcpic_path) != FileInfo_Get(apath):
            Doing(apath, username)
            Restart_Explorer()


if os.path.exists('Bing'):
    Bing_Doing()
elif os.path.exists('Focus'):
    Focus_Doing()
