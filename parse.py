import json
from numpy import less, tri
import openpyxl.cell
import requests
from importlib import reload
import time

import openpyxl
from openpyxl import load_workbook
from sympy import false
import config

def trim(string: str) -> str:
    if string == '': return ''
    try:
        while string[0]  == ' ': string = string[1:  ]
        while string[-1] == ' ': string = string[ :-1]
    except: pass
    return string

def update(group: str, course: str, debug: bool = false):
    files = json.load(open('files.json', 'r', encoding='utf-8'))
    url = files[group][course]

    response = requests.get(url, stream=True)

    print(f'\n{group} ({course})')  if debug else ...
    with open("temp/table.xlsx", "wb") as handle:
        for data in response.iter_content():
            handle.write(data)

    result = {}
    book = load_workbook('./temp/table.xlsx', data_only = True).active
    
    usedfrom = 0
    for i in range(1,30):
        try: d = book[f'A{i}'].value.replace(' ','')# type: ignore
        except: continue
        if d == 'ПОНЕДЕЛЬНИК':
            usedfrom = i-1
            print(f'\n\n\n{usedfrom=}\n\n')
            time.sleep(1)
            break
        
    
    for c in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        try: 
            typename = book[f'{c}{usedfrom}'].value.replace(' ','') # type: ignore
            if not typename: raise Exception
        except: continue
        if not typename: continue
        # if c != 'D': continue #test
        
        print(f'├── {typename}') if debug else ...
        result[typename] = {}

        last_day = ''
        last_delta = ''
        last_lesson = ''

        this = {}

        for r in range(usedfrom+1, 99):
            time.sleep(.01)
            day = book[f'A{r}'].value # type: ignore
            if day: day = day.capitalize()
            delta = book[f'B{r}'].value # type: ignore
            lesson = book[f'{c}{r}'].value # type: ignore

            print(day, delta, lesson, f'l_col:{c} row:{r}')  if debug else ...
            
            if day == 'Суббота': break

            if delta: 
                SS, EE = delta.split('-')[0], delta.split('-')[1]
                if len(SS) == 3: SS = f'{SS[0]}:{SS[1]}{SS[2]}'
                else: SS = f'{SS[0]}{SS[1]}:{SS[2]}{SS[3]}'

                if len(EE) == 3: EE = f'{EE[0]}:{EE[1]}{EE[2]}'
                else: EE = f'{EE[0]}{EE[1]}:{EE[2]}{EE[3]}'
                delta = f'{SS} - {EE}'

            print(f'{c}{r}', day, delta, 'Lesson text...' if lesson else 'None', book[f'{c}{r}']) if debug else ...  # type: ignore
            
            if day or delta or (lesson and trim(lesson) != '') or (not lesson and type(book[f'{c}{r}']) != openpyxl.cell.MergedCell): # type: ignore
                if day and day != last_day:
                    last_day = day
                    this[last_day] = {}

                if delta and delta != last_delta: 
                    last_delta = delta
                    this[last_day][last_delta] = []
                
                if lesson and trim(lesson) and lesson != last_lesson:
                    last_lesson = lesson
                    d = filter(lambda _: _!='', lesson.replace(',', '').replace('\n', '      ').split('   '))
                    lesson = []
                    for _ in d: lesson.append(trim(_))
                    lessonOBJ = {'type': '', 'name': '', 'info': [], 'raw': lesson, 'ok': True}
                    if lesson[0] == '*': 
                        lessonOBJ['ok'] = False
                        this[last_day][last_delta].append(lessonOBJ)
                        continue
                    
                    lessonOBJ['name'] = trim(lesson[0])
                    try: lessonOBJ['type'] = trim(lesson[1])
                    except:
                        lessonOBJ['ok'] = False
                        this[last_day][last_delta].append(lessonOBJ)
                        continue

                    for info in lesson[2:]:
                        if info == '' or info == ' ': continue
                        lessonOBJ['info'].append(trim(info))
                    this[last_day][last_delta].append(lessonOBJ)
                else: this[last_day][last_delta].append(None)



        result[typename] = this
    print(f'└──── Downloaded \n') if debug else ...
    json.dump(result, open(f'./timetables/{group} {course}.json'.replace(' ','_'), 'w', encoding='utf-8'), ensure_ascii=False, indent=4)
    return result


def updateAll(debug: bool = false) -> str: 
    global config
    config = reload(config)
    
    output = 'Result:'

    for group, courses in config.LOAD.items():
        output += f'\n* {group}: '
        for course in courses:
            try: r = update(group, course, debug)
            except Exception as e: 
                print(f'{group} > {course}: {e}') if debug else ...
                r = False
            output += f'{course} ' if r else f'[{course}] '

    return output
# updateAll()
