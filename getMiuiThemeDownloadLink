#获取小米主题商店主题主题直连

def getMiuiThemeDownloadLink(DownloadLink):
    import requests
    import json
    DownloadLink = DownloadLink.replace('zhuti.xiaomi.com/detail/','thm.market.xiaomi.com/thm/download/v2/')
    response = requests.get(DownloadLink)
    res = response.text
    res = json.loads(res)
    res = res['apiData']['downloadUrl']
    return res

DownloadLink = getMiuiThemeDownloadLink('http://zhuti.xiaomi.com/detail/0b72c1cf-2043-4b0d-97db-593b5258befb')
print (DownloadLink)
