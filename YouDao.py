#!/usr/bin/python
# -*- coding: UTF-8 -*-

import requests
from lxml import etree

def Translation(line):

    url = "http://m.youdao.com/translate"
    data = {"inputtext": line, "type": "AUTO"}
    con = requests.post(url, data).content
    html = etree.HTML(con)
    res = html.xpath("//ul[@id='translateResult']/li/text()")
    file_name = None
    if len(res) != 0:
        file_name = str(res[0]).replace(' ', '_')

    return file_name
