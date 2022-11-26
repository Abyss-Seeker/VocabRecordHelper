import requests
import json

# Global Variable Headers
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.81 Safari/537.36 Edg/104.0.1293.54'}

# Main body - searching word
def search(word):
    def org(res_dict):
        if 'entries' in res_dict['data']:
            explanation = res_dict['data']['entries'][0]
            if explanation['entry'] == word:
                definition = explanation['explain']
                return '; '.join(definition.split('；')[:3])
        return 'Word not found. Make sure it is in present tense and not spelled wrong! Try again :)'

    word = word.strip().lower()
    # 有道为动态加载，此链接可以从Network - XHR中找到，请求方式为requests
    url = 'https://dict.youdao.com/suggest?num=1&ver=3.0&doctype=json&cache=false&le=en&q={}'.format(word)
    res = requests.get(url, headers = headers)
    # 将获取的信息转为dict格式
    res_dict = res.json()
    definition = org(res_dict)
    return definition
    
