import os
import time

import zlib
import struct

import requests
from bs4 import BeautifulSoup as bs
import olefile
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import googlemaps
import re
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import json


def crawl(date):
    options = Options()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, "
                         "like Gecko) Chrome/77.0.3865.75 ""Safari/537.36")
    options.headless = True
    options.add_argument("disable-gpu")
    options.add_argument("disable-infobars")
    options.add_argument("--disable-extensions")
    prefs = {'profile.default_content_setting_values': {'cookies': 2, 'images': 2, 'plugins': 2, 'popups': 2,
                                                        'geolocation': 2, 'notifications': 2,
                                                        'auto_select_certificate': 2, 'fullscreen': 2, 'mouselock': 2,
                                                        'mixed_script': 2, 'media_stream': 2, 'media_stream_mic': 2,
                                                        'media_stream_camera': 2, 'protocol_handlers': 2,
                                                        'ppapi_broker': 2, 'automatic_downloads': 2, 'midi_sysex': 2,
                                                        'push_messaging': 2, 'ssl_cert_decisions': 2,
                                                        'metro_switch_to_desktop': 2, 'protected_media_identifier': 2,
                                                        'app_banner': 2, 'site_engagement': 2, 'durable_storage': 2},
             'download.default_directory': 'C:\projects\smartcity\data'}
    options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.implicitly_wait(3)

    url = "https://www.smpa.go.kr/user/nd54882.do"
    response = requests.get(url)

    html = response.text
    soup = bs(html, 'html.parser')

    for idx, sp in enumerate(soup.select(".subject > a")):
        if sp.text.strip().find(date) != -1:
            index = str(soup.select(".subject > a")[idx].attrs.values())[-14:-6]

    url = "https://www.smpa.go.kr/user/nd54882.do?View&uQ=&pageST=SUBJECT&pageSV=&imsi=imsi&page=1&pageSC=SORT_ORDER&pageSO=DESC&dmlType=&boardNo=" + index + "&returnUrl=https://www.smpa.go.kr:443/user/nd54882.do"

    driver.get(url)
    driver.find_element(By.CSS_SELECTOR, ".data-view > tbody > tr > td > a").click()
    time.sleep(0.2)

    day = "20" + date[:2] + "." + date[2:4] + "." + date[-2:]

    return day, os.listdir("C:\\projects\\smartcity\\data")[0]


def read(file_name):
    f = olefile.OleFileIO(file_name)
    dirs = f.listdir()

    header = f.openstream("FileHeader")
    header_data = header.read()
    is_compressed = (header_data[36] & 1) == 1

    nums = []
    for d in dirs:
        if d[0] == "BodyText":
            nums.append(int(d[1][len("Section"):]))
    sections = ["BodyText/Section" + str(x) for x in sorted(nums)]

    text = []
    for section in sections:
        bodytext = f.openstream(section)
        data = bodytext.read()
        if is_compressed:
            unpacked_data = zlib.decompress(data, -15)
        else:
            unpacked_data = data

        i = 0
        size = len(unpacked_data)
        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            rec_len = (header >> 20) & 0xfff

            if rec_type in [67]:
                rec_data = unpacked_data[i + 4:i + 2 + rec_len]
                decoded = rec_data.decode('utf-16')
                if decoded == "비고":
                    i += 4 + rec_len
                    break
            i += 4 + rec_len

        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            rec_len = (header >> 20) & 0xfff

            if rec_type in [67]:
                rec_data = unpacked_data[i + 4:i + 2 + rec_len]
                decoded = rec_data.decode('utf-16')
                text.append(decoded)

            i += 4 + rec_len

    f.close()

    return text


def make(listed_text):
    when = []
    where = []
    police = []
    police_str = ""
    where_cnt = -1
    police_cnt = -1

    for idx, text in enumerate(listed_text):
        if re.match("\d{2}:\d{2}∼\d{2}:\d{2}", listed_text[idx]) is not None:
            if police_str != "":
                police.append(police_str)
                police_str = ""

            when.append(text)
            where_cnt = idx + 1
        elif idx == where_cnt:
            where.append(text)
            police_cnt = idx + 3
        elif police_cnt == idx:
            if police_str != "":
                police_str += ", "
            police_str += text.replace(" ", "")
            police_cnt += 1

    if police_str != "":
        police.append(police_str)

    return when, where, police


def remove(when, where, police, place):
    dragon_when = []
    dragon_where = []
    dragon_police = []

    for idx, wheres in enumerate(where):
        if place != "":
            if police[idx].find(place) == -1:
                continue

        wheres = wheres.replace("", "")
        wheres = wheres.replace("∼", " ∼ ")
        wheres = wheres.replace("→", " → ")
        wheres = wheres.replace("⇄", " ⇄ ")
        wheres = wheres.replace("  ", " ")
        dragon_when.append(when[idx])
        dragon_where.append(wheres)
        dragon_police.append(police[idx])

    return dragon_when, dragon_where, dragon_police


def geoCode(where):
    gmaps = googlemaps.Client(key="AIzaSyAd7otLaukeX_N1yYu-l02OFeBP2xiAt6I")
    start = []
    end = []
    start_loc = []
    end_loc = []

    for place in where:
        re_place = re.findall("∼|→|⇄", place)
        if len(re_place) != 0:
            start.append(place.split(" ")[0])
            end.append(place.split(re_place[0])[1].split(" ")[1])
        else:
            start.append(place.split(" ")[0])
            end.append("")

    for idx in range(0, len(start)):
        try:
            start_result = gmaps.geocode(start[idx], language='ko')[0]
            start_loc.append(start_result['geometry']['location'])
        except:
            start_loc.append("Not find data")

        try:
            if end[idx] != "":
                end_result = gmaps.geocode(end[idx], language='ko')[0]
                end_loc.append(end_result['geometry']['location'])
            else:
                end_loc.append("")
        except:
            end_loc.append("Not find data")

    return start_loc, end_loc


def mk_json(day, when, where, police, start, end):
    day_list = []

    for idx in range(0, len(when)):
        day_dict = {"time": when[idx], "place": where[idx], "police": police[idx], "start": start[idx], "end": end[idx]}
        day_list.append(day_dict)

    json_day = {day: day_list}

    for i_path in os.listdir("C:\\projects\\smartcity\\data"):
        os.remove("C:\\projects\\smartcity\\data\\" + i_path)

    return json.dumps(json_day)
