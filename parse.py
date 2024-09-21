import json
import openpyxl.cell
import requests
from importlib import reload

import openpyxl
from openpyxl import load_workbook
import config


def update(group: str, course: str):
    files = json.load(open('files.json', 'r', encoding='utf-8'))
    url = files[group][course]

    response = requests.get(url, stream=True)

    print(f'\n{group} ({course})')
    with open("temp/table.xlsx", "wb") as handle:
        for data in response.iter_content():
            handle.write(data)

    result = {}
    book = load_workbook('./temp/table.xlsx', data_only = True).active
    for c in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        try: typename = book[f'{c}8'].value.replace(' ','')
        except: continue
        if not typename: continue
        # if c != 'D': continue #test
        
        print(f'├── {typename}')
        result[typename] = {}

        last_day = ''
        last_delta = ''
        last_lesson = ''

        this = {}

        for r in range(9, 99):
            day = book[f'A{r}'].value
            if day: day = day.capitalize()
            delta = book[f'B{r}'].value
            lesson = book[f'{c}{r}'].value
            
            if day == 'Суббота': break

            if delta: 
                SS, EE = delta.split('-')[0], delta.split('-')[1]
                if len(SS) == 3: SS = f'{SS[0]}:{SS[1]}{SS[2]}'
                else: SS = f'{SS[0]}{SS[1]}:{SS[2]}{SS[3]}'

                if len(EE) == 3: EE = f'{EE[0]}:{EE[1]}{EE[2]}'
                else: EE = f'{EE[0]}{EE[1]}:{EE[2]}{EE[3]}'
                delta = f'{SS} - {EE}'

            # print(f'{c}{r}', day, delta, 'Lesson text...' if lesson else 'None', book[f'{c}{r}'])
            
            if day or delta or (lesson) or (not lesson and type(book[f'{c}{r}']) != openpyxl.cell.MergedCell):
                if day and day != last_day:
                    last_day = day
                    this[last_day] = {}

                if delta and delta != last_delta: 
                    last_delta = delta
                    this[last_day][last_delta] = []
                
                if lesson and lesson != last_lesson:
                    
                    last_lesson = lesson
                    lessonOBJ = {'type': '', 'name': '', 'info': [], 'raw': lesson, 'ok': True}
                    if lesson[0] == '*': 
                        lessonOBJ['ok'] = False
                        this[last_day][last_delta].append(lessonOBJ)
                        continue
                    lessonOBJ['name'] = lesson.split('\n')[0]
                    try: lessonOBJ['type'] = lesson.split('\n')[1].split('  ')[0]
                    except:
                        lessonOBJ['ok'] = False
                        this[last_day][last_delta].append(lessonOBJ)
                        continue

                    addInfo = '  '.join(lesson.split('\n')[2:]).replace('  ', '$')
                    if len(lesson.split('\n')[1].split('  ')) > 1:
                        addInfo = '  '.join(lesson.split('\n')[1].split('  ')[1:]).replace('  ', '$') + '$$' + addInfo 
                    for info in addInfo.split('$'):
                        if info == '' or info == ' ': continue
                        while (info[0] == ' ' or info[-1] == ' '):
                            if info[0] == ' ': info = info[1:]
                            elif info[-1] == ' ': info = info[:-1]
                        lessonOBJ['info'].append(info)
                    this[last_day][last_delta].append(lessonOBJ)
                else: this[last_day][last_delta].append(None)



        result[typename] = this
    print(f'└──── Downloaded \n')
    json.dump(result, open(f'./timetables/{group} {course}.json'.replace(' ','_'), 'w', encoding='utf-8'), ensure_ascii=False, indent=4)



def updateAll(): 
    global config
    config = reload(config)

    for group, courses in config.LOAD.items():
        for course in courses:
            update(group, course)

updateAll()