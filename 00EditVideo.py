import time
import pyautogui as au
from pywinauto import application
import pandas as pd 
from EditName import EditName
from EditFileForm import EditFileForm
from GetPidList import GetPid
from PIL import ImageGrab



def moc(x, y):
    au.moveTo(x,y)
    au.click()


## Phase1_ 변수 설정##
# 참조파일 불러오기
df = pd.read_excel('./File/1차_합산본_20201215.xlsx', sheet_name='Sheet1')
df = df.dropna() # 빈셀 삭제


# 동영상 원본 파일 경로
tdir = 'C:/Users/pc1/Desktop/tmp/'
# 파일추가 
add = [2168, 95]
# 파일추가_탐색창
add_path = [1782, 316]
# 파일 이름
filename_input = [1641, 742]
# 곰믹스 워크스페이스 
gomwork = [2257, 165]
# 곰믹스 타임바 
gomtimebar = [168, 820]
# 분 선택
sel_min = [66, 684]
# 초 선택
sel_sec = [82, 684]
# 인코딩 시작
startbuttom = [2467, 1010]
# 인코딩 후 파일이름
savename = [1271, 620]
# 인코딩 완료
confirm = [1334, 821]
# 현재 프로젝트 
sel_p = [1974, 145]
# 최근 사용한 소스
sor = [1961, 172]
# 인코딩한 동영상
inco =[1976, 206]
# 삭제 버튼
b_del =[2522, 172]
# 프로그램 종료 시 저장여부 아니요 버튼
b_no = [1281, 587]


# app = application.Application()
# app.connect(process = GetPid("GomMixMain.exe")) #어플 실행 후 PID 자동으로 가져올 수 있도록 보완

## Phase2_ 편집 실행 ##

for i in range(200, len(df)):  #len(df)
    if i % 100 == 0: #100
        # 작업 프로그램 실행
        app = application.Application().start('C:/Program Files (x86)/GOM/GOMMix/GomMixMain.exe')
        time.sleep(7)
        # 작업파일 불러오기 
        fun, fin = EditName(i)
        au.moveTo(add[0], add[1])
        au.click()
        au.moveTo(add_path[0], add_path[1])
        au.click()
        au.typewrite(tdir+fin+'/MOVIE')
        au.press('enter')
        time.sleep(0.5)
        au.moveTo(filename_input[0], filename_input[1])
        au.click()
        au.typewrite(fun)
        au.press('enter')
        # 타임바에 소환
        au.moveTo(gomwork[0], gomwork[1])
        au.dragTo(gomtimebar[0], gomtimebar[1], 1, button='left')
        time.sleep(0.5)
        # 시간 편집
        au.moveTo(sel_min[0], sel_min[1]) #시작 시간    
        au.click()
        au.typewrite((str(df['ST'][i]).split(':')[0]),interval=0.1)    #분
        moc(sel_min[0], sel_min[1])
        au.typewrite((str(df['ST'][i]).split(':')[0]),interval=0.1) #자동조정 재설정
        au.moveTo(sel_sec[0], sel_sec[1])    
        au.click()    
        au.typewrite((str(df['ST'][i]).split(':')[1])+"50",interval=0.1)   #초 
        au.press('enter')
        au.hotkey('ctrl', 'x')
        au.moveTo(sel_min[0], sel_min[1]) #끝 시간
        au.click()
        au.typewrite(str(df['ET'][i]).split(':')[0],interval=0.1) #분
        moc(sel_min[0], sel_min[1])
        au.typewrite(str(df['ET'][i]).split(':')[0],interval=0.1) #자동조정 재설정
        au.moveTo(sel_sec[0], sel_sec[1])
        au.click()
        au.typewrite(str(df['ET'][i]).split(':')[1]+"50",interval=0.1) #초
        au.press('enter')
        au.hotkey('ctrl', 'x')
        au.press('delete')
        # au.press('left')
        au.press('left')
        time.sleep(0.5)
        au.press('delete')
        #인코딩 시작 
        au.moveTo(startbuttom[0], startbuttom[1])
        au.click()
        time.sleep(1)
        au.moveTo(savename[0], savename[1])
        au.click()
        au.hotkey('ctrl', 'a')
        au.press('delete')
        videotitle = "%d_%s_%d"%(int(df['Category'][i]), str(df['Name'][i][:-4]), int(df['Index'][i]))
        au.typewrite(videotitle)
        au.press('enter')
        # 인코딩 완료까지 대기
        while True:
            screen = ImageGrab.grab()
            if screen.getpixel((1476, 518)) == (217, 52, 66):
                time.sleep(1)
                break
            time.sleep(1) 
        au.moveTo(confirm[0], confirm[1])
        au.click()
        #잔여 파일 삭제
        au.moveTo(sel_p[0], sel_p[1])
        au.click()
        time.sleep(1)
        au.moveTo(gomwork[0], gomwork[1])
        au.click()
        au.press('delete')
        au.press('enter')
        au.moveTo(gomtimebar[0], gomtimebar[1])
        au.click()
        au.press('delete')
        moc(sor[0], sor[1])
        moc(b_del[0], b_del[1])
        au.press('enter')
        moc(inco[0], inco[1])
        moc(b_del[0], b_del[1])
        au.press('enter')
        moc(sel_p[0], sel_p[1])
        print('%s 개 완료'%i)
        time.sleep(1)
    else:
        # 작업파일 불러오기 
        fun, fin = EditName(i)
        au.moveTo(add[0], add[1])
        au.click()
        au.moveTo(add_path[0], add_path[1])
        au.click()
        au.typewrite(tdir+fin+'/MOVIE')
        au.press('enter')
        time.sleep(0.5)
        au.moveTo(filename_input[0], filename_input[1])
        au.click()
        au.typewrite(fun)
        au.press('enter')
        # 타임바에 소환
        au.moveTo(gomwork[0], gomwork[1])
        au.dragTo(gomtimebar[0], gomtimebar[1], 1, button='left')
        time.sleep(0.5)
        # 시간 편집
        au.moveTo(sel_min[0], sel_min[1]) #시작 시간    
        au.click()
        au.typewrite((str(df['ST'][i]).split(':')[0]),interval=0.1)    #분
        moc(sel_min[0], sel_min[1])
        au.typewrite((str(df['ST'][i]).split(':')[0]),interval=0.1) #자동조정 재설정
        au.moveTo(sel_sec[0], sel_sec[1])    
        au.click()    
        au.typewrite((str(df['ST'][i]).split(':')[1])+"50",interval=0.1)   #초 
        au.press('enter')
        au.hotkey('ctrl', 'x')
        au.moveTo(sel_min[0], sel_min[1]) #끝 시간
        au.click()
        au.typewrite(str(df['ET'][i]).split(':')[0],interval=0.1) #분
        moc(sel_min[0], sel_min[1])
        au.typewrite(str(df['ET'][i]).split(':')[0],interval=0.1) #자동조정 재설정
        au.moveTo(sel_sec[0], sel_sec[1])
        au.click()
        au.typewrite(str(df['ET'][i]).split(':')[1]+"50",interval=0.1) #초
        au.press('enter')
        au.hotkey('ctrl', 'x')
        au.press('delete')
        # au.press('left')
        au.press('left')
        time.sleep(0.5)
        au.press('delete')
        #인코딩 시작 
        au.moveTo(startbuttom[0], startbuttom[1])
        au.click()
        time.sleep(1)
        au.moveTo(savename[0], savename[1])
        au.click()
        au.hotkey('ctrl', 'a')
        au.press('delete')
        videotitle = "%d_%s_%d"%(int(df['Category'][i]), str(df['Name'][i][:-4]), int(df['Index'][i]))
        au.typewrite(videotitle)
        au.press('enter')
        # 인코딩 완료까지 대기
        while True:
            screen = ImageGrab.grab()
            if screen.getpixel((1476, 518)) == (217, 52, 66):
                time.sleep(1)
                break
            time.sleep(1) 
        au.moveTo(confirm[0], confirm[1])
        au.click()
        #잔여 파일 삭제
        au.moveTo(sel_p[0], sel_p[1])
        au.click()
        time.sleep(1)
        au.moveTo(gomwork[0], gomwork[1])
        au.click()
        au.press('delete')
        au.press('enter')
        au.moveTo(gomtimebar[0], gomtimebar[1])
        au.click()
        au.press('delete')
        moc(sor[0], sor[1])
        moc(b_del[0], b_del[1])
        au.press('enter')
        moc(inco[0], inco[1])
        moc(b_del[0], b_del[1])
        au.press('enter')
        moc(sel_p[0], sel_p[1])
        print('%s 개 완료'%i)
        time.sleep(1)
        if (i+1) % 100 == 0: #100
            au.hotkey("Alt", "F4")
            time.sleep(0.5)
            moc(b_no[0], b_no[1])
            time.sleep(1)
        else:
            pass

au.hotkey("Alt", "F4")
moc(b_no[0], b_no[1])