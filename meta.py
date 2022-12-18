import KartRider as KR
import os

#if not os.path.isdir("./metadata"):
#    KR.download_meta("./metadata")
#if os.path.isdir("./metadata"):
#    KR.set_metadatapath("./metadata")
def DownloadMeta(dir = "."):
    print("asdf")
    KR.download_meta(dir + "/metadata")
    print("fdsa")
api_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJYLUFwcC1SYXRlLUxpbWl0IjoiNTAwOjEwIiwiYWNjb3VudF9pZCI6IjE0OTM5MzMzMjYiLCJhdXRoX2lkIjoiMiIsImV4cCI6MTY4NDUwNTk5OSwiaWF0IjoxNjY4OTUzOTk5LCJuYmYiOjE2Njg5NTM5OTksInNlcnZpY2VfaWQiOiI0MzAwMTEzOTMiLCJ0b2tlbl90eXBlIjoiQWNjZXNzVG9rZW4ifQ.6iTKCqu7brkpuZaAlZmuN6MHiYc29VkPnnUTwJL92y0"

speed_team_list = ["a2bc792a460163af9b1d5726d8a125aa1ba1aa53b32a39be1bf0be2b7368154e"]
speed_indv_list = ["a2bc792a460163af9b1d5726d8a125aa1ba1aa53b32a39be1bf0be2b7368154e"]
item_team_list = [""]

member_accessId = {
    "감매": "1812190192",
    "구닌": "1963194707",
    "규현": "923309087",
    "그랑데": "201971733",
    "김준혁": "1963565777",
    "나노": "1493933326",
    "뚱쨕": "1242108075",
    "리즈": "84366047",
    "먼우금": "1493430569",
    "모나": "637808367",
    "밥줘": "520588857",
    "빠루": "1543956266",
    "소울": "520525983",
    "솖지": "722196376",
    "스코어": "1644475979",
    "영준": "956849994",
    "올렛": "1057161075",
    "이채팔": "1023634477",
    "인교": "2013630198",
    "자기야": "503843226",
    "재범": "402853456",
    "종이": "738847051",
    "쥐구쥐구": "251944965",
    "채우": "839373935",
    "쿠루루": "1980129461",
    "토박": "2080652850",
    "호영": "1577869385"
}