# uvicorn main:app --reload --host=0.0.0.0 --port=5000

import os

import requests

from models import model
from fastapi import FastAPI, Request

for path in os.listdir("C:\\projects\\smartcity\\data"):
    os.remove("C:\\projects\\smartcity\\data\\" + path)


def process(date, place):
    for i_path in os.listdir("C:\\projects\\smartcity\\data"):
        os.remove("C:\\projects\\smartcity\\data\\" + i_path)

    day, hwp_path = model.crawl(date)
    hwp_path = "C:\\projects\\smartcity\\data\\" + hwp_path
    listed_text = model.read(hwp_path)

    when, where, police = model.make(listed_text)

    when, where, police = model.remove(when, where, police, place)

    start, end = model.geoCode(where)

    day_json = model.mk_json(day, when, where, police, start, end)

    return day_json


app = FastAPI()


@app.get("/")
async def root():
    return {"message": "Hello World!!"}


@app.post("/work")
async def work(request: Request):
    request_data = await request.json()
    date, place = str(request_data['date']), str(request_data['where'])

    return process(date, place)
