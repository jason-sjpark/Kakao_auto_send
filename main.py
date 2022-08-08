from time import sleep
import win32con
import win32com.client
import win32api
import win32gui
import tkinter.font as tkFont
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
from tkinter import *

today_str = datetime.today().strftime("%Y-%m-%d")
today_date = datetime.today()

client_col = ['고객 연락처', '엔진오일 교체']

data = pd.read_csv('clients_data.csv')
# data = pd.read_csv('C:/자동발송/clients_data.csv')
# print(type(data))

# ls = [0 for _ in range(4)]

def make_gui():
    window = Tk()
    window.title('고객 자동 공지 프로그램')
    window.iconbitmap('icon.ico')
    # window.iconbitmap('C:/자동발송/icon.ico')
    window.geometry("800x500+510+290")
    window.resizable(False, False)
    fontStyle = tkFont.Font(family="MS Gothic", size=11)

    label = Label(window, text='고객 정보를 입력하세요.', font = fontStyle)
    label.pack()
    label.place(x=5, y=5)

    lbl_name = Label(window, text='고객 연락처: ', font = fontStyle)
    lbl_name.pack()
    lbl_name.place(x=5, y=40)

    global txt_name
    txt_name = Text(window, width = 15, height = 1, font = fontStyle)
    txt_name.pack()
    txt_name.place(x=100, y=40)
    # txt_name.get("1.0", END)   #1.0:  1(첫 번째 라인), 0(0번째 칼럼)
    # txt_name.delete("1.0", END)    # 첫 번째 라인부터 끝까지 삭제해라

    lbl_engine = Label(window, text = '엔진오일 교체', font = fontStyle)
    lbl_engine.pack()
    lbl_engine.place(x=240, y=40)

    engine_date = DateEntry(window, selectmode="day", locale='ko_KR', date_pattern='Y-mm-dd'
                            , year=today_date.year, month=today_date.month, day=today_date.day)
    engine_date._set_text(engine_date._date.strftime('%Y-%m-%d'))
    engine_date.pack()
    engine_date.bind("<<DateEntrySelected>>", date_entry_selected)
    engine_date.delete("0", END)    # 첫 번째 라인부터 끝까지 삭제해라
    engine_date.place(x=350, y=40)

    #고객 정보 추가하기
    btn_add = Button(window, width= 7, height = 1, padx= 10, pady = 10, text="추가하기", command = add, font = fontStyle)
    btn_add.pack()
    btn_add.place(x=600, y=30)

    #번호에 해당하는 고객 정보 조회
    btn_search = Button(window, width=7, height=1, padx=10, pady=10, text="조회하기", command=search, font = fontStyle)
    btn_search.pack()
    btn_search.place(x=600, y=80)

    # 전체 고객 정보 조회
    btn_all_search = Button(window, width=7, height=4, padx=10, pady=10, text="전체\n조회하기", command=all_search, font=fontStyle)
    btn_all_search.pack()
    btn_all_search.place(x=700, y=30)

    #초기화 하기
    btn_init = Button(window, width = 7, height = 1, text="초기화", command=init_data, font=fontStyle)
    btn_init.pack()
    btn_init.place(x=250, y=100)

    #알람 발송 대상 고객 조회
    btn_auto_search = Button(window, width=23, height=1, padx=10, pady=10, text="알람 발송 대상 고객 조회"
                             , command=send_search, font = fontStyle)
    btn_auto_search.pack()
    btn_auto_search.place(x=100, y=150)

    # 자동 발송하기
    btn_auto_send = Button(window, width=20, height=1, padx=10, pady=10, text="자동 발송하기"
                           , command=send, font = fontStyle)
    btn_auto_send.pack()
    btn_auto_send.place(x=400, y=150)

    #자동 발송 고객 리스트
    list_frame = Frame(window)
    list_frame.pack(fill="both", padx=5, pady=5)

    scrollbar = Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")
    listbox = Listbox(list_frame, width = 25, height = 15, selectmode="single", font = fontStyle, yscrollcommand=scrollbar.set)
    listbox.bind('<<ListboxSelect>>', onselect)
    listbox.insert(0, "text")

    listbox.pack(side="left", fill="both", expand=False)
    list_frame.place(x = 25, y = 210)
    scrollbar.config(command=listbox.yview)

    #자동 발송 메세지 내용
    txt_message = Text(window, width=40, height=15, font=fontStyle)
    txt_message.pack()
    txt_message.place(x=300, y=210)

    window.mainloop()


def date_entry_selected(event):
    w = event.widget
    global selected_date
    selected_date = w.get_date()
    print('Selected Date:{}'.format(selected_date))

def all_search():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel_file = excel.Workbooks.Open('C:/Users/GomZoo/PycharmProjects/ClientManagement/clients_data.csv')
    # excel_file = excel.Workbooks.Open('C:/자동발송/clients_data.csv')

def add():
    print(txt_name.get("1.0", END))  # 1.0:  1(첫 번째 라인), 0(0번째 칼럼)
    # df2 = pd.DataFrame([[5, 6], [7, 8]], columns=list('AB'), index=['x', 'y'])
    # df.append(df2)
    # data.append([txt_name.get("1.0", END)],  [selected_date])
    # data.to_csv('clients_data.csv', encoding='cp949')
    return

def search():
    return

def init_data():
    return

def send_search():
    return

def send():
    kakao_control()
    return

def onselect(evt):
    w = evt.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    print(value)

def kakao_control():
    open_chatroom(kakao_opentalk_name) # 채팅방 열기
    text = contents
    kakao_sendtext(kakao_opentalk_name, text) # 메시지 전송
    sleep(1)
    close_chatroom(kakao_opentalk_name)

# 채팅방 열기
def open_chatroom(chatroom_name):
    # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)

    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx(hwndkakao_edit1, None, "EVA_Window", None)  # 친구창에서 검색
    # hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)    #채팅목록에서 검색
    hwndkakao_edit3 = win32gui.FindWindowEx(hwndkakao_edit2_1, None, "Edit", None)

    # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    sleep(1)  # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    sleep(1)

#엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

#채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    #핸들 채팅방
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None)
    #hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

def close_chatroom(kakao_opentalk_name):
    #닫으면서 날짜 수정 해야함(다음 교체 날짜로)
    hwndkakao = win32gui.FindWindow(None, kakao_opentalk_name)
    win32api.SendMessage(hwndkakao, win32con.WM_CLOSE, 0, 0)

if __name__ == '__main__':
    make_gui()