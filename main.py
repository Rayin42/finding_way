import time, win32con, win32api, win32gui, ctypes
import pyperclip
import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler


#채팅방 이름
kakao_open_chat = '샆'
discord_open_chat = "떠돌이｜상인 - Discord"
#검색할 단어
command1 = '전설'
command2 = '웨이'
#최신 데이터를 다루기 위해 이전 기록을 저장
global cls

#특수키 입력 등을 위한 세팅
PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput
MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW
MakeLong = win32api.MAKELONG
w = win32con

#카카오톡이 활성되어 있을 때 채팅방 검색, 그 창을 띄움
#핸들을 찾아갈 때 친구검색 -> 채팅방 검색 순으로 이루어져야 하기 때문에 edit2_1 -> edit2_2순으로 해야할 필요가 있었음
def open_chat(chat_name):
    hwnd_kakao = win32gui.FindWindow(None, "카카오톡")
    hwnd_kakao_edit1 = win32gui.FindWindowEx(hwnd_kakao, None, "EVA_ChildWindow", None)
    hwnd_kakao_edit2_1 = win32gui.FindWindowEx(hwnd_kakao_edit1, None, "EVA_Window", None)
    hwnd_kakao_edit2_2 = win32gui.FindWindowEx(hwnd_kakao_edit1, hwnd_kakao_edit2_1, "EVA_Window", None)
    hwnd_kakao_edit3 = win32gui.FindWindowEx(hwnd_kakao_edit2_2, None, "Edit", None)

    win32api.SendMessage(hwnd_kakao_edit3, win32con.WM_SETTEXT, 0, chat_name)
    time.sleep(1)
    send_return(hwnd_kakao_edit3)
    time.sleep(1)


#카카오톡 채팅방 -> 글자입력 파트를 찾아 메시지를 입력
def send_kakaotext(chat_name, text):
    hwnd_main = win32gui.FindWindow(None, chat_name)
    hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RICHEDIT50W", None)

    win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, text)
    send_return(hwnd_edit)


#엔터 역할
def send_return(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def copy_chat(chat_name):
    #오류 방지용 클립보드 내용 덮어쓰기
    pyperclip.copy(" ")
    #디스코드 이름으로 윈도우 핸들을 찾아내 ctrl A, ctrl C로 모든 대화내용 복사
    hwnd_main = win32gui.FindWindow(None, chat_name)
    PostKeyEx(hwnd_main, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    PostKeyEx(hwnd_main, ord('C'), [w.VK_CONTROL], False)

    #테스트 시 test_case
    # test = "test_case"
    # hwnd_main = win32gui.FindWindow(None, test)
    # PostKeyEx(hwnd_main, ord('A'), [w.VK_CONTROL], False)
    # time.sleep(1)
    # PostKeyEx(hwnd_main, ord('C'), [w.VK_CONTROL], False)

    #클립보드 내용 붙여넣기
    ctext = pyperclip.paste()
    return ctext


def PostKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):
        ThreadId = GetWindowThreadProcessId(hwnd, None)
        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()
            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128
            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)
        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)


#프로그램을 구동했을 때 이미 진행된 대화를 cls에 저장
def save_last_chat():
    count = 0
    while True:
        ttext = copy_chat(discord_open_chat)
        a = ttext.split('\r\n')
        df = pd.DataFrame(a)
        if len(df) > 3:
            break
        count += 1
        if count > 5:
            return -1
    print(df)
    return df.iloc[-2, 0]


def check_chat_command(cls):
    while True:
        ttext = copy_chat(discord_open_chat)
        a = ttext.split('\r\n')
        df = pd.DataFrame(a)
        if len(df) > 3:
            break

    #로그 확인용
    print(f"{time.localtime().tm_year}-{time.localtime().tm_mon}-{time.localtime().tm_mday} / "
          f"{time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}")

    #기존 저장되어 있던 데이터와 비교, 잘라내서 최신 데이터만 확인할 수 있도록
    #df_index = df.index[df[0] == cls]
    for i in range(len(df)):
        if df.iloc[-i, 0] == cls:
            df_index = len(df)-i
            break

    df1 = df.iloc[df_index + 1:, 0]

    # #지정한 커맨드가 잘라낸 데이터에 속해 있는지 파악, 1개 이상 있을 경우 카카오톡을 열고 순서대로 발송
    found = df1[df1.str.contains(command1)]
    if int(found.count()) > 0:
        open_chat(kakao_open_chat)
        for i in range(found.count()):
            send_kakaotext(kakao_open_chat, found.iloc[i])
            time.sleep(0.5)
        print("전호!")

    #커맨드 2번 있는지 파악(여유 되면 나중에 데코레이터를 쓰든 합치든)
    found = df1[df1.str.contains(command2)]
    if int(found.count()) > 0:
        open_chat(kakao_open_chat)
        for i in range(found.count()):
            send_kakaotext(kakao_open_chat, found.iloc[i])
            time.sleep(0.5)
        print("웨이!")

    return df.iloc[-2, 0]


#스케쥴러 수행명령
def job_1(clss):
    global cls
    cls = check_chat_command(clss)


def main():
    global cls
    cls = save_last_chat()
    if cls == -1:
        print("디스코드 행방불명")
        return -1
    print("대응 디스코드 방 : " + discord_open_chat)
    print("대응 카카오톡 방 : " + kakao_open_chat)
    print("init copy complete")

    #스케쥴러 사용 시 parameter 변수가 갱신이 되지 않음 -> 시계 구현
    # sched = BackgroundScheduler(timezone="utc")
    # sched.start()
    # sched.add_job(job_1, 'cron', minute="40", second='01', id="test_01", args=[cls])
    # sched.add_job(job_1, 'cron', minute="45", second='01', id="test_02", args=[cls])
    #sched.add_job(job_1, 'interval', seconds=3, id="test_03", args=[cls])
    count = 0
    while True:
        time.sleep(1)
        count += 1
        if count == 60:
            count = 0
            if time.localtime().tm_min == 35:
                job_1(cls)
            if time.localtime().tm_min == 40:
                job_1(cls)
            if time.localtime().tm_min == 42:
                job_1(cls)
            elif time.localtime().tm_min == 45:
                job_1(cls)


if __name__ == '__main__':
    main()

