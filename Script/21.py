import re
import requests
import json



url = 'https://manager.starday.shop/?#/turnuserecord'




# 获取请求url域名
def getHost(url):
    pattern = re.compile(r'(.*?)://(.*?)/', re.S)
    response = re.search(pattern, url)
    if response:
        return {'header': str(response.group(1)).strip(), 'host': str(response.group(2)).strip()}
    else:
        return None


# 创建存储cookie的map
cookieType = {}


# 获取cookie
def getCode():
    locationList = set()
    cookie = ''

    resUrl = 'https://manager.starday.shop/?#/turnuserecord'  # 开始重定的url
    header = {'User-Agen': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'}
    try:
        while True:
            # 每次请求重定向中的url    allow_redirects=False 禁止重定向
            response = requests.get(url=resUrl, headers=header, allow_redirects=False)
            #cookie = response.cookies.get_dict()
            hostObj = getHost(resUrl)
            if hostObj == {} :
                return None
            cookie = response.cookies.get_dict()
            if cookie == {}:
                pass
            else:
                cookieType[str(hostObj['host']).strip()] = json.dumps(cookie)  # 保存cookie
            # 获取跳转的url
            if 'Location' in response.headers.keys():
                url = response.headers.get('Location')
                if not 'http' in url:
                    url = hostObj['header'] + '://' + hostObj['host'] + url  # 拼接host域名
                resUrl = url
                if url in locationList: break
                locationList.add(url)
            else:
                break;
        return {'url': str(resUrl), 'content': response.content, 'header': response.headers}
    except TypeError:
        return "";
    except UnboundLocalError:
        return None;





getHost(url)
getCode()
###NT