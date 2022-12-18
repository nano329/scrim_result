from collections import Counter
from main import SetTime, api, record, avg
import xlsxwriter

arr = [(0, 0), (0, 1), (0, 2), (0, 3), (1, 0), (1, 1), (1, 2), (1, 3)]

def team_search(nick, start, end, team1, team2):
    start = SetTime(start)
    end = SetTime(end)
    print(".")
    speed_match = api.user(nick).getMatches(limit=50, match_types="스피드 팀전", start_date=start, end_date=end)
    print(",")
    item_match_get = api.user(nick).getMatches(limit=50, start_date=start, end_date=end)
    print(item_match_get)
    item_match = {}
    if 'Unknown' in item_match_get.keys(): item_match['Unknown'] = item_match_get['Unknown']
    if '아이템 팀전' in item_match_get.keys(): item_match['아이템 팀전'] = item_match_get['아이템 팀전']
    print(".")



    print(speed_match)
    print(item_match)
    speed_rank = [[[], [], [], []], [[], [], [], []]]
    speed_kart = [[[], [], [], []], [[], [], [], []]]
    speed_main = [[], []]
    speed_track = []
    speed_win = []
    speed_record = [[[], [], [], []], [[], [], [], []]]
    item_rank = [[[], [], [], []], [[], [], [], []]]
    item_kart = [[[], [], [], []], [[], [], [], []]]
    item_main = [[], []]
    item_track = []
    item_win = []
    item_record = [[[], [], [], []], [[], [], [], []]]
    for gametype, match in speed_match.items():
        for game in match:
            print(len(game.detail.teams[0]))
            if len(game.detail.teams[0]) != 4:
                continue
            for i, j in arr:
                player = game.detail.teams[i][j]
                if len(speed_rank[i][j]) == 0:
                    speed_rank[i][j].insert(1, player.charactername)
                speed_rank[i][j].insert(1, player.matchrank)
                speed_kart[i][j].insert(0, player.kart)
                speed_record[i][j].insert(0, player.matchtime)
                print(".", end='')
            speed_track.insert(0, game.detail.track)
            speed_win.insert(0, game.detail.matchresult)
    print("asdfasdfasdf")

    for gametype, match in item_match.items():
        for game in match:
            print(len(game.detail.teams[0]))
            if len(game.detail.teams[0]) != 4:
                continue
            for i, j in arr:
                player = game.detail.teams[i][j]
                if len(item_rank[i][j]) == 0:
                    item_rank[i][j].insert(1, player.charactername)
                item_rank[i][j].insert(1, player.matchrank)
                item_kart[i][j].insert(0, player.kart)
                item_record[i][j].insert(0, player.matchtime)

            item_track.insert(0, game.detail.track)
            item_win.insert(0, game.detail.matchresult)

    speed_1st = []
    item_1st = []
    print(speed_rank)
    print(item_rank)
    print(speed_win)
    print(speed_kart)
    print(item_kart)
    print(speed_record)
    print(item_record)

    for i in range(len(speed_record[0][0])):
        lst = []
        for j in range(2):
            for k in range(4):
                if speed_record[j][k][i] != 0:
                    lst.append(speed_record[j][k][i])
        speed_1st.append(min(lst))
    for i in range(len(item_record[0][0])):
        lst = []
        for j in range(2):
            for k in range(4):
                if item_record[j][k][i] != 0:
                    lst.append(item_record[j][k][i])
        item_1st.append(min(lst))
    print(item_1st)

    for i in range(2):
        for j in range(4):
            if len(speed_track) > 0:
                speed_main[i].append((Counter(speed_kart[i][j]).most_common(1)[0][0]))
            if len(item_track) > 0:
                item_main[i].append((Counter(item_kart[i][j]).most_common(1)[0][0]))

    print(speed_main)
    print(item_main)

    for i in range(4):
        if speed_rank[1][i][0] == nick:
            team1, team2 = team2, team1

    speed_info = {
        "len" : len(speed_track),
        "match" : speed_match,
        "rank" : speed_rank,
        "kart" : speed_kart,
        "main" : speed_main,
        "track" : speed_track,
        "win" : speed_win,
        "record" : speed_record,
        "1st" : speed_1st
    }

    item_info = {
        "len" : len(item_track),
        "match": item_match,
        "rank": item_rank,
        "kart": item_kart,
        "main": item_main,
        "track": item_track,
        "win": item_win,
        "record": item_record,
        "1st": item_1st
    }

    date = start.split()[0]

    return speed_info, item_info


def team_write(speed_info, item_info, team1, team2, date, filename, dir=".", ret=9):
    wb = xlsxwriter.Workbook(f'{dir}/{filename}.xlsx')

    speed_match = speed_info["match"]
    speed_rank = speed_info["rank"]
    speed_kart = speed_info["kart"]
    speed_main = speed_info["main"]
    speed_track = speed_info["track"]
    speed_win = speed_info["win"]
    speed_record = speed_info["record"]
    speed_1st = speed_info["1st"]

    item_match = item_info["match"]
    item_rank = item_info["rank"]
    item_kart = item_info["kart"]
    item_main = item_info["main"]
    item_track = item_info["track"]
    item_win = item_info["win"]
    item_record = item_info["record"]
    item_1st = item_info["1st"]


    dark_format = wb.add_format(
        {"bg_color": "#000000", "font_color": "#ffffff", "font_size": 9, "align": "center", "valign": "center"})
    title = wb.add_format(
        {"bg_color": "#000000", "font_color": "#ffffff", "font_size": 24, "align": "center", "valign": "center"})
    red_team = wb.add_format(
        {"bg_color": "#c00000", "font_color": "#ffffff", "font_size": 16, "align": "center", "valign": "center"})
    blue_team = wb.add_format(
        {"bg_color": "#002060", "font_color": "#ffffff", "font_size": 16, "align": "center", "valign": "center"})
    red_win = wb.add_format(
        {"bg_color": "#f4b084", "font_color": "#000000", "font_size": 12, "align": "center", "valign": "center"})
    red_1st = wb.add_format(
        {"bg_color": "#ed7d31", "font_color": "#ffffff", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})
    red_win_r = wb.add_format(
        {"bg_color": "#f4b084", "font_color": "#ff0000", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})
    red_track = wb.add_format(
        {"bg_color": "#ed7d31", "font_color": "#000000", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})
    blue_win = wb.add_format(
        {"bg_color": "#8ea9db", "font_color": "#000000", "font_size": 12, "align": "center", "valign": "center"})
    blue_win_r = wb.add_format(
        {"bg_color": "#8ea9db", "font_color": "#ff0000", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})
    blue_1st = wb.add_format(
        {"bg_color": "#4472c4", "font_color": "#ffffff", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})
    blue_track = wb.add_format(
        {"bg_color": "#4472c4", "font_color": "#000000", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})
    lose = wb.add_format(
        {"bg_color": "#808080", "font_color": "#000000", "font_size": 12, "align": "center", "valign": "center"})
    lose_r = wb.add_format(
        {"bg_color": "#808080", "font_color": "#f00000", "font_size": 12, "bold": True, "align": "center",
         "valign": "center"})

    print("allsllelll", f'{dir}/{filename}.xlsx')

    if len(speed_track):
        ws = wb.add_worksheet(f'{filename} SPEED')
        ws.set_column('A:H', 2)
        ws.set_column('I:I', 45)
        ws.set_column('J:Q', 2)
        for i in "BDFHKMOQ":
            ws.set_column(i + ':' + i, 12)

        ws.merge_range('A1:H1', team1, blue_team)
        ws.write('I1', date, title)
        ws.write('I2', 'SPEED', dark_format)
        ws.merge_range('J1:Q1', team2, red_team)
        for i in [2, len(speed_rank[0][0]) + 2, len(speed_rank[0][0]) + 3]:
            for j in [65, 67, 69, 71, 74, 76, 78, 80]:
                ws.merge_range(chr(j) + str(i) + ':' + chr(j + 1) + str(i), '', dark_format)
        for i, j in arr:
            player_list = speed_rank[i][j]
            ws.write(1, i * 9 + j * 2, player_list[0], dark_format)
            for m in range(1, len(player_list)):
                apply_format = None
                if str(i + 1) != speed_win[m - 1]:  # win
                    if i == 0:
                        apply_format = blue_1st if player_list[m] == 1 else blue_win_r if player_list[
                                                                                              m] == -1 else blue_win
                    else:
                        apply_format = red_1st if player_list[m] == 1 else red_win_r if player_list[
                                                                                            m] == -1 else red_win
                else:
                    apply_format = lose_r if player_list[m] == -1 else lose
                ws.write(m + 1, i * 9 + j * 2, 9 if player_list[m] == -1 else player_list[m], apply_format)
                apply_format.set_font_size(8)
                ws.write(m + 1, i * 9 + j * 2 + 1, speed_kart[i][j][m - 1], apply_format)

            ws.write(len(player_list) + 1, i * 9 + j * 2, speed_main[i][j], dark_format)
            ws.write(len(player_list) + 2, i * 9 + j * 2, str(int(avg(player_list[1:], ret)) / 100) + '등', dark_format)
        ws.write(len(speed_rank[0][0]) + 1, 8, "메인 카트", dark_format)
        ws.write(len(speed_rank[0][0]) + 2, 8, "평균 순위", dark_format)
        for i in range(len(speed_track)):
            apply_format = red_track if speed_win[i] == '1' else blue_track
            ws.write(i + 2, 8, speed_track[i] + f'({record(speed_1st[i])})', apply_format)

    if len(item_track):
        ws = wb.add_worksheet(f'{filename} ITEM')
        ws.set_column('A:L', 2)
        ws.set_column('M:M', 45)
        ws.set_column('N:Y', 2)
        for i in "BEHKORUX":
            ws.set_column(i + ':' + i, 12)
        for i in "CFILPSVY":
            ws.set_column(i + ':' + i, 6)

        ws.merge_range('A1:L1', team1, blue_team)
        ws.write('M1', date, title)
        ws.write('M2', 'ITEM', dark_format)
        ws.merge_range('N1:Y1', team2, red_team)
        for i in [2, len(item_rank[0][0]) + 2]:
            for j in [65, 68, 71, 74, 78, 81, 84, 87]:
                ws.merge_range(chr(j) + str(i) + ':' + chr(j + 2) + str(i), '', dark_format)
        for i, j in arr:
            player_list = item_rank[i][j]
            player_kart = item_kart[i][j]
            print(player_list)
            print(player_kart)
            print(len(player_list), len(item_1st))
            ws.write(1, i * 13 + j * 3, player_list[0], dark_format)
            # ws["abcdfghi"[i * 4 + j] + "1"] = player_list[0]
            for m in range(1, len(player_list)):
                apply_format = None
                if str(i + 1) != item_win[m - 1]:  # win
                    if i == 0:
                        apply_format = blue_1st if player_list[m] == 1 else blue_win
                    else:
                        apply_format = red_1st if player_list[m] == 1 else red_win
                else:
                    apply_format = lose
                apply_format.set_font_size(8)
                ws.write(m + 1, i * 13 + j * 3, 9 if player_list[m] == -1 else player_list[m], apply_format)
                ws.write(m + 1, i * 13 + j * 3 + 1, player_kart[m - 1], apply_format)
                ws.write(m + 1, i * 13 + j * 3 + 2,
                         f'({"Retire" if item_record[i][j][m - 1] == 0 else "+" + str(record(item_record[i][j][m - 1] - item_1st[m - 1]))})',
                         apply_format)
            ws.write(len(player_list) + 1, i * 12 + j * 3, item_main[i][j], dark_format)
        ws.write(len(item_rank[0][0]) + 1, 12, "메인 카트", dark_format)

        for i in range(len(item_track)):
            ws.write(i + 2, 12, item_track[i] + f'({record(item_1st[i])})',
                     red_track if item_win[i] == '1' else blue_track)

    wb.close()