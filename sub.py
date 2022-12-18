import requests
from bs4 import BeautifulSoup
import pandas as pd

API_Key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhY2NvdW50X2lkIjoiMTQ5MzkzMzMyNiIsImF1dGhfaWQiOiIyIiwidG9rZW5fdHlwZSI6IkFjY2Vzc1Rva2VuIiwic2VydmljZV9pZCI6IjQzMDAxMTM5MyIsIlgtQXBwLVJhdGUtTGltaXQiOiI1MDA6MTAiLCJuYmYiOjE2NTMzMDI2MDEsImV4cCI6MTY2ODg1NDYwMSwiaWF0IjoxNjUzMzAyNjAxfQ.72nr5PzFwWSI3c_14xNguGY3DEvy-zA86kWnF1VQrR4"
URL = "https://api.nexon.co.kr/kart/v1.0/matches/{match_id}".format(API= API_KEY)

# API로 데이터 불러오기
rq = requests.get(URL)
soup = BeautifulSoup(rq.text, "html.parser")
[출처] [크롤링] 파이썬(Python)으로 공공데이터 오픈 API 사용하기|작성자 Mr WOO

