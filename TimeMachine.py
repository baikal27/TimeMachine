# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gibeum.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import numpy as np
import glob, os
import re
import pandas as pd
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import QAbstractTableModel, Qt
from matplotlib import pyplot as plt, rcParams, font_manager

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 640)
        font = QtGui.QFont()
        font.setPointSize(22)
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btn_upload = QtWidgets.QPushButton(self.centralwidget)
        self.btn_upload.setGeometry(QtCore.QRect(665, 130, 130, 50))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_upload.setFont(font)
        self.btn_upload.setObjectName("btn_upload")
        self.select_month = QtWidgets.QComboBox(self.centralwidget)
        self.select_month.setGeometry(QtCore.QRect(140, 40, 104, 26))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.select_month.setFont(font)
        self.select_month.setEditable(True)
        self.select_month.setDuplicatesEnabled(True)
        self.select_month.setObjectName("select_month")
        self.select_name = QtWidgets.QComboBox(self.centralwidget)
        self.select_name.setGeometry(QtCore.QRect(420, 40, 104, 26))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.select_name.setFont(font)
        self.select_name.setEditable(True)
        self.select_name.setDuplicatesEnabled(True)
        self.select_name.setObjectName("select_name")
        self.lbl_start = QtWidgets.QLabel(self.centralwidget)
        self.lbl_start.setGeometry(QtCore.QRect(110, 17, 160, 20))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.lbl_start.setFont(font)
        self.lbl_start.setAlignment(QtCore.Qt.AlignCenter)
        self.lbl_start.setObjectName("lbl_start")
        self.lbl_end = QtWidgets.QLabel(self.centralwidget)
        self.lbl_end.setGeometry(QtCore.QRect(390, 17, 160, 20))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.lbl_end.setFont(font)
        self.lbl_end.setAlignment(QtCore.Qt.AlignCenter)
        self.lbl_end.setObjectName("lbl_end")
        self.tableview = QtWidgets.QTableView(self.centralwidget)
        self.tableview.setGeometry(QtCore.QRect(15, 70, 643, 471))
        self.tableview.setObjectName("tableview")
        self.btn_save_excel = QtWidgets.QPushButton(self.centralwidget)
        self.btn_save_excel.setGeometry(QtCore.QRect(665, 190, 130, 50))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_save_excel.setFont(font)
        self.btn_save_excel.setObjectName("btn_save_excel")
        self.btn_path = QtWidgets.QPushButton(self.centralwidget)
        self.btn_path.setGeometry(QtCore.QRect(665, 70, 130, 50))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_path.setFont(font)
        self.btn_path.setObjectName("btn_path")
        self.btn_preview = QtWidgets.QPushButton(self.centralwidget)
        self.btn_preview.setGeometry(QtCore.QRect(190, 545, 150, 45))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_preview.setFont(font)
        self.btn_preview.setObjectName("btn_preview")
        self.btn_postview = QtWidgets.QPushButton(self.centralwidget)
        self.btn_postview.setGeometry(QtCore.QRect(350, 545, 150, 45))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_postview.setFont(font)
        self.btn_postview.setObjectName("btn_postview")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(16)
        self.statusbar.setFont(font)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        # 한글설정
        #rcParams['font.sans-serif'] = 'Source Han Sans K'
        #rcParams['font.weight'] = 'regular'
        #rcParams['axes.titlesize'] = 15
        #rcParams['ytick.labelsize'] = 12
        #rcParams['xtick.labelsize'] = 12

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        mon_items = [str(i) + '월' for i in range(1, 13)]
        self.select_month.addItems(mon_items)
        name_items = sorted(['안성진', '정종섭', '박형민', '이효동', '김다운', '김정은', '김태우', '신용철', '오최근',
                      '유지현', '이상록', '정상은', '정신우', '최윤호', '하지성', '김소희', '박미소', '장현동',
                      '정훈', '최누리', '김범수', '하원혁', '김연수', '박솔빛나', '백형권', '최봉규', '천기범',
                      '이분녀', '이채연', '유병규', '유준형', '금윤수', '김선희', '김현종', '오갑순', '임영우', '정재환'])
        name_items.insert(0, '전체')
        self.select_name.addItems(name_items)

        self.btn_path.clicked.connect(self.find_path)
        self.btn_upload.clicked.connect(self.go_upload)
        self.btn_save_excel.clicked.connect(self.save_excel)
        self.btn_preview.clicked.connect(self.preview)
        self.btn_postview.clicked.connect(self.postview)
        self.texting = 'save the time then you will be free \n'
        self.statusbar.showMessage(self.texting)

        self.code = False

        global standard_inouttime
        standard_inouttime = 5 * 60 * 60 + 30 * 60  # 5:30

        self.base_dir = os.getcwd()
        self.filepath = self.base_dir

    def find_path(self):
        # filepath, _ = QFileDialog.getOpenFileName(MainWindow, "working_dir", os.getcwd(), "csv (*.csv)")
        self.filepath = QFileDialog.getExistingDirectory(MainWindow, caption='Select directory',
                                                         directory=self.base_dir)
        if self.filepath:
            os.chdir(self.filepath)
            self.texting = '파일 경로를 {}로 변경합니다.'.format(self.filepath)
            self.statusbar.showMessage(self.texting)
        else:
            os.chdir(self.base_dir)
            self.texting = 'PATH를 눌러 파일 경로를 새로 지정해 주세요. 아니면 {}가 default 경로로 설정됩니다.'.format(self.base_dir)
            self.statusbar.showMessage(self.texting)

    def go_upload(self):
        df = pd.DataFrame([])  # 빈 df를 넘겨서 테이블을 없앨려고 하는 것임. 계속 이 모델 방식을 써야할 듯.
        model = MyTableModel(df)
        self.tableview.setModel(model)
        self.tableview.show()

        self.month = self.select_month.currentText()
        self.name = self.select_name.currentText()
        self.filename = []

        erp = glob.glob('*.xlsx')
        file1 = self.month + ' erp신청.xlsx'
        file2 = self.month + ' 출입기록.xlsx'

        if file1 in erp and file2 in erp:
            self.filename.append(file1)
            self.filename.append(file2)
            self.code = True
            self.do_uploading()
        else:
            self.code = False
            self.texting = 'There is no input file for analysis. Please check PATH.'
            self.statusbar.showMessage(self.texting)

    def do_uploading(self):

        self.texting = '{}의 {} 시간외근무근무 보상시간 계산을 시작합니다. \n'.format(self.name, self.month)
        self.statusbar.showMessage(self.texting)

        sheetname = 'erp 신청데이터'
        appdata = pd.DataFrame([])
        col_list = []

        reque_col_name = ['부서', '직급', '직위', '신청이름', '날짜', '연장휴일', '신청시작', '신청종료', '휴게시간']
        reque_data = pd.read_excel(self.filename[0], encoding='cp949', names=reque_col_name,
                                   usecols=[0, 1, 2, 3, 4, 6, 7, 8, 12])

        reque_data[['신청시작hour', '신청시작min']] = reque_data['신청시작'].str.split(':', n=2, expand=True).astype(np.int64)
        reque_data[['신청종료hour', '신청종료min']] = reque_data['신청종료'].str.split(':', n=2, expand=True).astype(np.int64)
        reque_data[['휴게시간hour', '휴게시간min']] = reque_data['휴게시간'].str.split(':', n=2, expand=True).astype(np.int64)
        reque_data['reque_start_total_secs'] = reque_data.apply(lambda x: 60 * 60 * x['신청시작hour'] + 60 * x['신청시작min'],
                                                                axis=1)
        reque_data['reque_end_total_secs'] = reque_data.apply(lambda x: 60 * 60 * x['신청종료hour'] + 60 * x['신청종료min'],
                                                              axis=1)
        reque_data['reque_rest_total_secs'] = reque_data.apply(lambda x: 60 * 60 * x['휴게시간hour'] + 60 * x['휴게시간min'],
                                                               axis=1)
        reque_data['AB_groups'] = reque_data.apply(lambda x: self.det_groups(x['신청이름']), axis=1)

        # ---------------------------------- 입출입 자료 처리 -----------------------------------------------------
        inout_col_name = ['발생시각', '출입이름', '출입직급']
        inout_data = pd.read_excel(self.filename[1], encoding='cp949', names=inout_col_name, usecols=[1, 5, 9])
        inout_data[['발생날짜', '발생시각2']] = inout_data['발생시각'].str.split(' ', n=2, expand=True)
        inout_data[['년', '월', '일']] = inout_data['발생날짜'].str.split('-', n=3, expand=True)
        inout_data[['년', '월', '일']] = inout_data[['년', '월', '일']].astype(np.int64)
        inout_data[['시', '분', '초']] = inout_data['발생시각2'].str.split(':', n=3, expand=True)
        inout_data[['시', '분', '초']] = inout_data[['시', '분', '초']].astype(np.int64)
        inout_data['total_secs'] = inout_data.apply(lambda x: 60 * 60 * x['시'] + 60 * x['분'] + x['초'], axis=1)

        inout_data['new일'] = inout_data.apply(lambda x: self.day(x['일'], x['total_secs']), axis=1)
        inout_data['new시'] = inout_data.apply(lambda x: self.tim(x['시'], x['total_secs']), axis=1)
        inout_data['new발생날짜'] = inout_data.apply(lambda x: self.makedate(x['년'], x['월'], x['new일']), axis=1)
        inout_data['new_total_secs'] = inout_data.apply(lambda x: 60 * 60 * x['new시'] + 60 * x['분'] + x['초'], axis=1)

        new_inout_data = inout_data.loc[:,
                         ['발생시각', 'new발생날짜', '출입이름', '출입직급', '년', '월', 'new일', 'new시', '분', '초', 'new_total_secs']]
        # print(new_inout_data.loc[1150:1200])

        # name = '하지성'
        if self.name == '전체':
            name_list = [i for i in set(new_inout_data['출입이름'])]
        else:
            name_list = [self.name]

        dummy_list = []
        for name in name_list:
            teher = new_inout_data.loc[new_inout_data['출입이름'] == name]
            teher_list = [i for i in set(teher['new발생날짜'])]
            for i in teher_list:
                teher2 = teher.loc[teher['new발생날짜'] == i]
                # print(teher2)
                # print('idx_max: {}'.format(teher2['new_total_secs'].idxmax(axis=0)))
                # print('idx_min" {}'.format(teher2['new_total_secs'].idxmin(axis=0)))
                idx_max = teher2['new_total_secs'].idxmax(axis=0)
                idx_min = teher2['new_total_secs'].idxmin(axis=0)
                if len(teher2.index) == 1:
                    inhour = teher2.loc[teher2.index[0], 'new시']
                    inmin = teher2.loc[teher2.index[0], '분']
                    insec = teher2.loc[teher2.index[0], '초']
                    intotalsecs = teher2.loc[teher2.index[0], 'new_total_secs']
                    outhour = inhour
                    outmin = inmin
                    outsec = insec
                    outtotalsecs = intotalsecs
                # print('min : {} {} {}'.format(inhour, inmin, insec))
                # print('max : {} {} {}'.format(outhour, outmin, outsec))
                elif len(teher2.index) >= 2:
                    inhour = teher2.loc[idx_min, 'new시']
                    inmin = teher2.loc[idx_min, '분']
                    insec = teher2.loc[idx_min, '초']
                    intotalsecs = teher2.loc[idx_min, 'new_total_secs']
                    outhour = teher2.loc[idx_max, 'new시']
                    outmin = teher2.loc[idx_max, '분']
                    outsec = teher2.loc[idx_max, '초']
                    outtotalsecs = teher2.loc[idx_max, 'new_total_secs']
                # print('min : {} {} {}'.format(inhour, inmin, insec))
                # print('max : {} {} {}'.format(outhour, outmin, outsec))
                # print(new_inout_data.loc[1150:1200])
                new_inout_data2 = [teher2.loc[idx_max, '발생시각'], i,
                                   teher2.loc[idx_max, '출입이름'], teher2.loc[idx_max, '출입직급'],
                                   teher2.loc[idx_max, '년'], teher2.loc[idx_max, '월'], teher2.loc[idx_max, 'new일'],
                                   inhour, inmin, insec, outhour, outmin, outsec, intotalsecs, outtotalsecs]
                # final_inout_col_names = ['io발생시각', 'io발생날짜', 'io출입이름', 'io출입직급', 'io년', 'io월', 'io일', 'inhour', 'inmin', 'insec', 'outhour', 'outmin', 'outsec', 'intotalsecs', 'outtotalsecs']
                # print(new_inout_data2_col_names)
                # print(new_inout_data2)
                dummy_list.append(new_inout_data2)
        # print(dummy_list)
        final_inout_col_names = ['io발생시각', 'io날짜', 'io이름', 'io출입직급', 'io년', 'io월', 'io일',
                                 'inhour', 'inmin', 'insec', 'outhour', 'outmin', 'outsec',
                                 'intotalsecs', 'outtotalsecs']
        final_inout_data = pd.DataFrame(dummy_list, columns=final_inout_col_names)

        # ------------------------ add two dataframes -----------------------------------------------------
        #
        # reque_data_col_name = ['부서', '직급', '직위', '신청이름', '날짜', '연장휴일', '신청시작', '신청종료', '신청시작hour',
        #                  '신청시작min','신청종료hour', '신청종료min', '휴게시간hour', '휴게시간min',
        #                  'reque_start_total_secs', 'reque_end_total_secs', 'reque_rest_total_secs', 'AB_groups']
        # final_inout_data_col_names = ['io발생시각', 'io날짜', 'io이름', 'io출입직급', 'io년', 'io월', 'io일', 'inhour',
        #                              'inmin', 'insec', 'outhour', 'outmin', 'outsec', 'intotalsecs', 'outtotalsecs']

        if self.name == '전체':
            self.reque_name_list = [i for i in set(reque_data['신청이름'])]
        else:
            self.reque_name_list = [self.name]
        dummy_list = []

        for reque_name in self.reque_name_list:
            reque_small = reque_data.loc[reque_data['신청이름'] == reque_name]
            final_inout_small = final_inout_data.loc[final_inout_data['io이름'] == reque_name]
            # reque_date_list = [i for i in set(reque_small['날짜'])]
            reque_date_list = [i for i in set(reque_small['날짜'])]
            # print(reque_small.dtypes)
            # print(reque_small[['신청이름', '날짜']])
            # print(final_inout_small[['io이름', 'io날짜']])

            for reque_date in reque_date_list:
                reque_indexraw = reque_small.loc[reque_small["날짜"] == reque_date].index
                final_index = final_inout_small.loc[final_inout_small["io날짜"] == reque_date].index
                final_index = final_index[0]

                for i in range(len(reque_indexraw)):
                    reque_index = reque_indexraw[i]
                    # print(final_index)
                    # print(reque_small.loc[reque_index, ["신청이름", "날짜"]])
                    # print(final_inout_small.loc[final_index, ["io이름", "io날짜"]])
                    # print('happy')

                    reque_depart = reque_small.loc[reque_index, "부서"]
                    reque_grade = reque_small.loc[reque_index, "직급"]
                    reque_position = reque_small.loc[reque_index, "직위"]
                    reque_holiday = reque_small.loc[reque_index, "연장휴일"]
                    reque_name = reque_small.loc[reque_index, "신청이름"]
                    reque_date = reque_small.loc[reque_index, "날짜"]
                    reque_group = reque_small.loc[reque_index, "AB_groups"]
                    # reque_start_time = ':'.join([reque_small.loc[reque_index, '신청시작hour'].astype(np.str), reque_small.loc[reque_index, '신청시작min'].astype(np.str)])
                    reque_start_time = str(reque_small.loc[reque_index, '신청시작hour']) + ':' + str(
                        reque_small.loc[reque_index, '신청시작min'])
                    # reque_end_time = ':'.join(reque_small.loc[reque_index, '신청종료hour'], reque_small.loc[reque_index, '신청종료min'])
                    reque_end_time = str(reque_small.loc[reque_index, '신청종료hour']) + ':' + str(
                        reque_small.loc[reque_index, '신청종료min'])
                    reque_start_secs = reque_small.loc[reque_index, "reque_start_total_secs"]
                    reque_end_secs = reque_small.loc[reque_index, "reque_end_total_secs"]
                    reque_rest_secs = reque_small.loc[reque_index, "reque_rest_total_secs"]

                    inout_date = final_inout_small.loc[final_index, 'io날짜']
                    inout_name = final_inout_small.loc[final_index, 'io이름']
                    inout_intime = str(final_inout_small.loc[final_index, 'inhour']) + ':' + \
                                   str(final_inout_small.loc[final_index, 'inmin']) + ':' + \
                                   str(final_inout_small.loc[final_index, 'insec'])
                    inout_outtime = str(final_inout_small.loc[final_index, 'outhour']) + ':' + \
                                    str(final_inout_small.loc[final_index, 'outmin']) + ':' + \
                                    str(final_inout_small.loc[final_index, 'outsec'])
                    inout_start_secs = final_inout_small.loc[final_index, 'intotalsecs']
                    inout_end_secs = final_inout_small.loc[final_index, 'outtotalsecs']
                    # print(reque_name, reque_start, inout_start)

                    if inout_start_secs == inout_end_secs:
                        if inout_end_secs <= reque_start_secs:
                            final_start = inout_start_secs
                            final_end = inout_end_secs
                        elif reque_start_secs < inout_end_secs <= reque_end_secs:
                            final_start = reque_start_secs
                            final_end = inout_end_secs
                        elif reque_start_secs < reque_end_secs <= inout_end_secs:
                            final_start = reque_start_secs
                            final_end = reque_end_secs
                        else:
                            print('입출입시간과 요청시간 사이에 또 어떤 case가 있을까? 일단 근무시간 No 인정으로 간다.')
                            final_start = inout_start_secs
                            final_end = inout_start_secs
                    elif inout_start_secs < inout_end_secs:
                        if inout_start_secs < reque_start_secs <= inout_end_secs:
                            if reque_start_secs < reque_end_secs <= inout_end_secs:
                                final_start = reque_start_secs
                                final_end = reque_end_secs
                            elif inout_end_secs < reque_end_secs:
                                final_start = reque_start_secs
                                final_end = inout_end_secs
                            else:
                                print('요청시간 시작과 끝이 잘못된 경우. final_start, final_end 모두 0으로 처리.')
                                fina_start = 0
                                final_end = 0
                        elif reque_start_secs <= inout_start_secs:
                            if inout_end_secs < reque_end_secs:
                                final_start = inout_start_secs
                                final_end = inout_end_secs
                            elif inout_start_secs < reque_end_secs <= inout_end_secs:
                                final_start = inout_start_secs
                                final_end = reque_end_secs
                            elif reque_end_secs <= inout_start_secs:
                                final_start = reque_start_secs
                                final_end = reque_end_secs
                                print('출근 지문을 안찍었다고 판단함.')
                            else:
                                print('요청시간 시작과 끝이 출근시간보다 빠른 경우. final_start, final_end 모두 0으로 처리.')
                                fina_start = 0
                                final_end = 0
                        else:
                            print("요청시간이 퇴근시간 이후인 경우. final_start, final_end 모두 0으로 처리.")
                            final_start = 0
                            final_end = 0
                    else:
                        final_start = 0
                        final_end = 0

                    dummy_list.append([reque_name, reque_date, reque_group, reque_holiday, reque_grade, reque_depart,
                                       reque_position, reque_start_time, reque_end_time, reque_start_secs,
                                       reque_end_secs, reque_rest_secs, inout_date, inout_name, inout_intime,
                                       inout_outtime, inout_start_secs, inout_end_secs, final_start, final_end])
        # print(dummy_list)

        self.super_final_col_names = ["reque_name", "reque_date", "reque_group", "reque_holiday", "reque_grade",
                                      "reque_depart", "reque_position", "reque_start_time", "reque_end_time",
                                      "reque_start_secs", "reque_end_secs", "reque_rest_secs",
                                      "inout_date", "inout_name", "inout_intime", "inout_outtime",
                                      "inout_start_secs", "inout_end_secs", "final_start", "final_end"]
        self.super_final_data = pd.DataFrame(dummy_list, columns=self.super_final_col_names)
        print(self.super_final_data.loc[:, ['reque_date', 'reque_start_time', 'reque_end_time', 'reque_rest_secs',
                                            'inout_intime', 'inout_outtime', 'final_start', 'final_end']].sort_values(
            'reque_date').to_string())
        print(self.super_final_data.loc[:,
              ['reque_date', 'reque_holiday', 'final_start', 'final_end', 'reque_rest_secs']].to_string())

        # ---------------------------------------- 보상시간 산출 ---------------------------------------------
        # -------------------------------------------------------------------------------------------------
        standard_time = 18 * 60 * 60 + 20 * 60  # 18:20
        self.super_final_data['rewardsecs'] = self.super_final_data.apply(
            lambda x: self.calcul_reward(x['reque_holiday'], x['final_start'], x['final_end'], x['reque_rest_secs']),
            axis=1)
        self.super_final_data['hms_final_start'] = self.super_final_data.apply(
            lambda x: self.secs_to_hms(x['final_start']), axis=1)
        self.super_final_data['hms_final_end'] = self.super_final_data.apply(lambda x: self.secs_to_hms(x['final_end']),
                                                                             axis=1)
        self.super_final_data['hms_rewardsecs'] = self.super_final_data.apply(
            lambda x: self.secs_to_hms(x['rewardsecs']), axis=1)
        # super_final_col_list = ["reque_name", "reque_date", "reque_group", "reque_holiday", "reque_grade",
        #                         "reque_depart", "reque_position", "reque_start_time", "reque_end_time",
        #                         "reque_start_secs", "reque_end_secs", "reque_rest_secs",
        #                         "inout_date", "inout_name", "inout_intime", "inout_outtime",
        #                         "inout_start_secs", "inout_end_secs", "final_start", "final_end",
        #                         'rewardsecs', 'hms_final_start', 'hms_final_end', 'hms_rewardsecs']

        self.daily_final_data = self.super_final_data.loc[:,
                                ['reque_depart', 'reque_name', 'reque_group', 'reque_grade',
                                 'reque_date', 'reque_holiday', 'reque_start_time', 'reque_end_time',
                                 'inout_intime', 'inout_outtime', 'hms_final_start', 'hms_final_end',
                                 'reque_rest_secs', 'rewardsecs', 'hms_rewardsecs']]
        #print(self.daily_final_data[
        #          ['reque_date', 'reque_holiday', 'reque_rest_secs', 'rewardsecs', 'hms_rewardsecs']].sort_values(
        #    'reque_date').to_string)
        recolumns = {'reque_depart':'부서', 'reque_name':'이름', 'reque_group':'AB그룹', 'reque_grade':'직급',
                     'reque_date':'신청날짜', 'reque_holiday':'근무형태', 'reque_start_time':'신청시작시간',
                     'reque_end_time':'신청종료시간', 'inout_intime':'출근시간', 'inout_outtime':'퇴근시간',
                     'hms_final_start':'결정시작시간', 'hms_final_end':'결정종료시간', 'reque_rest_secs':'휴게시간',
                     'rewardsecs':'보상시간', 'hms_rewardsecs':'보상시간(hms)'}
        self.report_daily = self.daily_final_data.rename(columns=recolumns)
        self.report2_daily = self.report_daily.loc[:,['부서', '이름', '신청날짜', '근무형태', '신청시작시간', '신청종료시간', '출근시간', '퇴근시간',
                                     '결정시작시간', '결정종료시간', '보상시간(hms)']]
        print(self.report2_daily.set_index(['이름', '신청날짜']).sort_index())

        model = MyTableModel(self.report2_daily)
        self.tableview.setModel(model)
        self.tableview.show()

        dummy_list = []
        for name in self.reque_name_list:
            indexed = self.daily_final_data.loc[self.daily_final_data['reque_name'] == name, 'rewardsecs'].index[0]
            # print(indexed)

            monthly_depart = self.daily_final_data.loc[indexed, 'reque_depart']
            monthly_group = self.daily_final_data.loc[indexed, 'reque_group']
            monthly_name = self.daily_final_data.loc[indexed, 'reque_name']
            monthly_grade = self.daily_final_data.loc[indexed, 'reque_grade']

            total_rewardsecs = self.daily_final_data.loc[self.daily_final_data['reque_name'] == name, 'rewardsecs'].sum(
                axis=0)
            hms_total_rewardsecs = self.secs_to_hms(total_rewardsecs)
            dummy_list.append(
                [monthly_depart, monthly_name, monthly_group, monthly_grade, total_rewardsecs, hms_total_rewardsecs])

        monthly_final_data_col_names = ['monthly_depart', 'monthly_name', 'monthly_group', 'monthly_grade',
                                        'total_rewardsecs', 'hms_total_rewardsecs']
        monthly_final_data = pd.DataFrame(dummy_list, columns=monthly_final_data_col_names)

        monthly_final_data['hms_money_reward'] = monthly_final_data.apply(
            lambda x: self.money_reward(x['total_rewardsecs'], x['monthly_grade'], x['monthly_group']), axis=1)
        monthly_final_data['hms_vacation_reward'] = monthly_final_data.apply(
            lambda x: self.vacation_reward(x['total_rewardsecs'], x['monthly_grade'], x['monthly_group']), axis=1)

        recolumns = {'monthly_depart':'부서', 'monthly_name':'이름', 'monthly_group':'AB그룹', 'monthly_grade':'직급',
                     'total_rewardsecs':'총보상시간', 'hms_total_rewardsecs':'총보상시간(hms)',
                     'hms_money_reward':'금전보상시간', 'hms_vacation_reward':'휴가보상시간'}
        self.report_monthly = monthly_final_data.rename(columns=recolumns)
        print(self.report_monthly.set_index(['부서', '이름']).sort_index())

        print(self.report_monthly.columns)
        self.texting = '{}의 {} 시간외근무근무 보상시간을 성공적으로 계산하였습니다. \n'.format(self.name, self.month)
        self.statusbar.showMessage(self.texting)

    def save_excel(self):
        result_dir_name = 'result_excel'
        if os.path.exists(result_dir_name) is False:
            os.mkdir(result_dir_name)

        if self.code:
            self.excel_dir = os.path.join(self.filepath, result_dir_name)
            print(self.excel_dir)
            re_name_daily = '_'.join([self.month, self.name, 'daily.xlsx'])
            self.report2_daily.to_excel(os.path.join(self.excel_dir, re_name_daily), sheet_name='daily')

            re_name_monthly = '_'.join([self.month, self.name, 'monthly.xlsx'])
            self.report_monthly.to_excel(os.path.join(self.excel_dir, re_name_monthly), sheet_name='monthly')
            # columns = ["부서", "직급", "이름", "총보상시간", "금전보상시간", "휴가보상시간"])
            self.texting = '{}의 {} 시간외근무 보상시간을 일별({}), 월 합계({}) 파일로 {}에 저장하였습니다. \n'.format(self.name, self.month, re_name_daily, re_name_monthly, self.excel_dir)
            self.statusbar.showMessage(self.texting)
        else:
            self.texting = '저장할 엑셀파일이 없습니다. PATH와 입력자료를 다시 확인하세요. \n'
            self.statusbar.showMessage(self.texting)

    def preview(self):
        if self.code:
            #model = MyTableModel(self.report2_daily.set_index(['이름', '신청날짜']).sort_index())
            model = MyTableModel(self.report2_daily)
            self.tableview.setModel(model)
            self.tableview.show()
            self.texting = '{}의 {} 일별 환산 결과를 보여줍니다.\n'.format(self.name, self.month)
            self.statusbar.showMessage(self.texting)
        else:
            self.texting = '보여줄 일별 환산 자료가 없습니다. PATH와 DATA UPLOAD를 다시 확인하세요.\n'
            self.statusbar.showMessage(self.texting)


    def postview(self):
        if self.code:
            #model = MyTableModel(self.report_monthly.set_index(['부서', '이름']).sort_index())
            model = MyTableModel(self.report_monthly)
            self.tableview.setModel(model)
            self.tableview.show()
            self.texting = '{}의 {} 월 합산 결과를 보여줍니다.\n'.format(self.name, self.month)
            self.statusbar.showMessage(self.texting)
        else:
            self.texting = '보여줄 월 합산 자료가 없습니다. PATH와 DATA UPLOAD를 다시 확인하세요. \n'
            self.statusbar.showMessage(self.texting)

    def day(self, day, total_secs):
        global standard_inouttime
        if total_secs < standard_inouttime:
            return day - 1
        else:
            return day

    def tim(self, time, total_secs):
        global standard_inouttime
        if total_secs < standard_inouttime:
            return time + 24
        else:
            return time

    def makedate(self, year, mon, day):
        mon = str(mon)
        if len(mon) == 1:
            mon = '0' + mon
        else:
            mon = mon
        day = str(day)
        if len(day) == 1:
            day = '0' + day
        else:
            day = day
        newdate = '-'.join([str(year), mon, day])
        return newdate

    def det_groups(self, name):
        A_group = ['천기범', '최봉규', '이채연']
        if name in A_group:
            return 'A'
        else:
            return 'B'

    def secs_to_hms(self, secs):
        hour = secs // 3600
        rest_sec = secs % 3600
        min = rest_sec // 60
        rest_sec2 = rest_sec % 60
        return '{}:{}:{}'.format(int(hour), int(min), int(rest_sec2))

    def calcul_reward(self, norholy, start_secs, end_secs, rest_secs):
        standard_secs = 22 * 60 * 60  # 보상시간이 2배로 뛰는 기준 시간
        stand_4_hour = 4 * 60 * 60
        stand_8_hour = 8 * 60 * 60
        stand_12_hour = 12 * 60 * 60
        diff = end_secs - start_secs
        if norholy == '휴일근무':
            if rest_secs == 0:
                if diff <= stand_4_hour:
                    if start_secs <= end_secs <= standard_secs:
                        rewardsecs = diff * 1.5
                    elif start_secs <= standard_secs <= end_secs:
                        rewardsecs = (standard_secs - start_secs) * 1.5 + (end_secs - standard_secs) * 2.0
                    elif standard_secs <= start_secs <= end_secs:
                        rewardsecs = diff * 2.0
                    else:
                        print("I dont know what case1 is")
                elif stand_4_hour < diff <= stand_8_hour:
                    rest = 30 * 60
                    diff -= rest
                    chpoint = start_secs + rest
                    if start_secs <= end_secs <= standard_secs:
                        rewardsecs = diff * 1.5
                    elif start_secs <= standard_secs <= end_secs:
                        rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                    elif standard_secs <= start_secs <= end_secs:
                        rewardsecs = diff * 2.0
                    else:
                        print("I dont know what case2 is")
                elif stand_8_hour < diff:
                    rest = 60 * 60
                    diff -= rest
                    if stand_8_hour <= diff:
                        if start_secs < end_secs <= standard_secs:
                            diff_before = stand_8_hour
                            diff_after = diff - stand_8_hour
                            rewardsecs = diff_before * 1.5 + diff_after * 2.0
                        elif start_secs < standard_secs < end_secs:
                            chpoint = start_secs + stand_8_hour
                            rewardsecs = (standard_secs - start_secs) * 1.5 + (chpoint - standard_secs) * 2.0 + \
                                         (end_secs - chpoint) * 2.5
                        elif standard_secs <= start_secs < end_secs:
                            chpoint = start_secs + stand_8_hour
                            rewardsecs = (chpoint - start_secs) * 2.0 + (end_secs - chpoint) * 2.5
                        else:
                            print("I dont know what case3 is")
                    else:
                        if start_secs < end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs < standard_secs < end_secs:
                            rewardsecs = (standard_secs - start_secs) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs < end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case4 is")
            else:
                if 0 < rest_secs <= 30 * 60:
                    if diff <= stand_4_hour:
                        diff -= rest_secs
                        chpoint = start_secs + rest_secs
                        if start_secs <= end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs <= standard_secs <= end_secs:
                            rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs <= end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case5 is")
                    elif stand_4_hour < diff <= stand_8_hour:
                        rest = 30 * 60
                        diff -= rest
                        chpoint = start_secs + rest
                        if start_secs <= end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs <= standard_secs <= end_secs:
                            rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs <= end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case6 is")
                    elif stand_8_hour < diff:
                        rest = 1 * 60 * 60
                        diff -= rest
                        if stand_8_hour <= diff:
                            if start_secs < end_secs <= standard_secs:
                                diff_before = stand_8_hour
                                diff_after = diff - stand_8_hour
                                rewardsecs = diff_before * 1.5 + diff_after * 2.0
                            elif start_secs < standard_secs < end_secs:
                                chpoint = start_secs + stand_8_hour
                                rewardsecs = (standard_secs - start_secs) * 1.5 + (chpoint - standard_secs) * 2.0 + \
                                             (end_secs - chpoint) * 2.5
                            elif standard_secs <= start_secs < end_secs:
                                chpoint = start_secs + stand_8_hour
                                rewardsecs = (chpoint - start_secs) * 2.0 + (end_secs - chpoint) * 2.5
                            else:
                                print("I dont know what case7 is")
                        else:
                            if start_secs < end_secs <= standard_secs:
                                rewardsecs = diff * 1.5
                            elif start_secs < standard_secs < end_secs:
                                rewardsecs = (standard_secs - start_secs) * 1.5 + (end_secs - standard_secs) * 2.0
                            elif standard_secs <= start_secs < end_secs:
                                rewardsecs = diff * 2.0
                            else:
                                print("I dont know what case8 is")
                elif 30 * 60 < rest_secs <= 1 * 60 * 60:
                    if diff <= stand_4_hour:
                        diff -= rest_secs
                        chpoint = start_secs + rest_secs
                        if start_secs <= end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs <= standard_secs <= end_secs:
                            rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs <= end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case9 is")
                    elif stand_4_hour < diff <= stand_8_hour:
                        diff -= rest_secs
                        chpoint = start_secs + rest_secs
                        if start_secs <= end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs <= standard_secs <= end_secs:
                            rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs <= end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case10 is")
                    elif stand_8_hour < diff:
                        rest = 1 * 60 * 60
                        diff -= rest
                        if stand_8_hour <= diff:
                            if start_secs < end_secs <= standard_secs:
                                diff_before = stand_8_hour
                                diff_after = diff - stand_8_hour
                                rewardsecs = diff_before * 1.5 + diff_after * 2.0
                            elif start_secs < standard_secs < end_secs:
                                chpoint = start_secs + stand_8_hour
                                rewardsecs = (standard_secs - start_secs) * 1.5 + (chpoint - standard_secs) * 2.0 + \
                                             (end_secs - chpoint) * 2.5
                            elif standard_secs <= start_secs < end_secs:
                                chpoint = start_secs + stand_8_hour
                                rewardsecs = (chpoint - start_secs) * 2.0 + (end_secs - chpoint) * 2.5
                            else:
                                print("I dont know what case11 is")
                        else:
                            if start_secs < end_secs <= standard_secs:
                                rewardsecs = diff * 1.5
                            elif start_secs < standard_secs < end_secs:
                                rewardsecs = (standard_secs - start_secs) * 1.5 + (end_secs - standard_secs) * 2.0
                            elif standard_secs <= start_secs < end_secs:
                                rewardsecs = diff * 2.0
                            else:
                                print("I dont know what case12 is")
                elif 1 * 30 * 60 < rest_secs:
                    if diff <= stand_4_hour:
                        diff -= rest_secs
                        chpoint = start_secs + rest_secs
                        if start_secs <= end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs <= standard_secs <= end_secs:
                            rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs <= end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case13 is")
                    elif stand_4_hour < diff <= stand_8_hour:
                        diff -= rest_secs
                        chpoint = start_secs + rest_secs
                        if start_secs <= end_secs <= standard_secs:
                            rewardsecs = diff * 1.5
                        elif start_secs <= standard_secs <= end_secs:
                            rewardsecs = (standard_secs - chpoint) * 1.5 + (end_secs - standard_secs) * 2.0
                        elif standard_secs <= start_secs <= end_secs:
                            rewardsecs = diff * 2.0
                        else:
                            print("I dont know what case14 is")
                    elif stand_8_hour < diff:
                        diff -= rest_secs
                        if stand_8_hour <= diff:
                            if start_secs < end_secs <= standard_secs:
                                diff_before = stand_8_hour
                                diff_after = diff - stand_8_hour
                                rewardsecs = diff_before * 1.5 + diff_after * 2.0
                            elif start_secs < standard_secs < end_secs:
                                chpoint = start_secs + stand_8_hour
                                rewardsecs = (standard_secs - start_secs) * 1.5 + (chpoint - standard_secs) * 2.0 + \
                                             (end_secs - chpoint) * 2.5
                            elif standard_secs <= start_secs < end_secs:
                                chpoint = start_secs + stand_8_hour
                                rewardsecs = (chpoint - start_secs) * 2.0 + (end_secs - chpoint) * 2.5
                            else:
                                print("I dont know what case15 is")
                        else:
                            if start_secs < end_secs <= standard_secs:
                                rewardsecs = diff * 1.5
                            elif start_secs < standard_secs <= end_secs:
                                rewardsecs = (standard_secs - start_secs) * 1.5 + (end_secs - standard_secs) * 2.0
                            elif standard_secs < start_secs <= end_secs:
                                rewardsecs = diff * 2.0
                            else:
                                print("I dont know what case16 is")
        else:
            if start_secs >= 18 * 60 * 60 + 20 * 60:  # 시간외 근무는 18:20 이후부터 올린다는 암묵적 합의

                if start_secs < standard_secs <= end_secs:
                    rewardsecs = ((standard_secs - start_secs) * 1.5) + ((end_secs - standard_secs) * 2.0)
                elif start_secs < end_secs < standard_secs:
                    rewardsecs = (end_secs - start_secs) * 1.5
                elif standard_secs <= start_secs < end_secs:
                    rewardsecs = (end_secs - start_secs) * 2.0
                else:
                    print(
                        "function of normal_day_evening has calculated wrong rewardsecs. It should not gonna be happened!")
            elif start_secs == 18 * 60 * 60:
                if 0 < rest_secs:
                    start_secs += rest_secs
                    if start_secs < standard_secs <= end_secs:
                        rewardsecs = ((standard_secs - start_secs) * 1.5) + ((end_secs - standard_secs) * 2.0)
                    elif start_secs < end_secs < standard_secs:
                        rewardsecs = (end_secs - start_secs) * 1.5
                    elif standard_secs <= start_secs < end_secs:
                        rewardsecs = (end_secs - start_secs) * 2.0
                    else:
                        print(
                            "function of normal_day_evening has calculated wrong rewardsecs. It should not gonna be happened!")
                else:
                    start_secs += 30 * 60
                    if start_secs < standard_secs <= end_secs:
                        rewardsecs = ((standard_secs - start_secs) * 1.5) + ((end_secs - standard_secs) * 2.0)
                    elif start_secs < end_secs < standard_secs:
                        rewardsecs = (end_secs - start_secs) * 1.5
                    elif standard_secs <= start_secs < end_secs:
                        rewardsecs = (end_secs - start_secs) * 2.0
                    else:
                        print(
                            "function of normal_day_evening has calculated wrong rewardsecs. It should not gonna be happened!")

            else:
                rewardsecs = (end_secs - start_secs) * 1.5
            # return rewardsecs
        return rewardsecs

    def money_reward(self, total_secs, grade, group):
        alpha_A = {'2급': 3 * 60 * 60, '3급': 9 * 60 * 60, '4급': 24 * 60 * 60, '5급': 39 * 60 * 60,
                   '6급': 43 * 60 * 60 + 30 * 60,
                   '가급': 9 * 60 * 60, '나급': 12 * 60 * 60, '다급': 21 * 60 * 60, '라급': 34 * 60 * 60 + 30 * 60,
                   '마급': 43 * 60 * 60 + 30 * 60, '바급': 45 * 60 * 60}
        alpha_B = {'2급': 3 * 60 * 60, '3급': 15 * 60 * 60, '4급': 33 * 60 * 60, '5급': 48 * 60 * 60,
                   '6급': 52 * 60 * 60 + 30 * 60,
                   '가급': 12 * 60 * 60, '나급': 21 * 60 * 60, '다급': 34 * 60 * 60 + 30 * 60, '라급': 43 * 60 * 60 + 50 * 60,
                   '마급': 52 * 60 * 60 + 30 * 60, '바급': 63 * 60 * 60}
        if group == 'A':
            if total_secs <= alpha_A[grade]:
                money_secs = total_secs
                vacation_secs = 0
            else:
                money_secs = alpha_A[grade]
                vacation_secs = total_secs - alpha_A[grade]
        else:
            if total_secs <= alpha_B[grade]:
                money_secs = total_secs
                vacation_secs = 0
            else:
                money_secs = alpha_B[grade]
                vacation_secs = total_secs - alpha_B[grade]
        return self.secs_to_hms(money_secs)

    def vacation_reward(self, total_secs, grade, group):
        alpha_A = {'2급': 3 * 60 * 60, '3급': 9 * 60 * 60, '4급': 24 * 60 * 60, '5급': 39 * 60 * 60,
                   '6급': 43 * 60 * 60 + 30 * 60,
                   '가급': 9 * 60 * 60, '나급': 12 * 60 * 60, '다급': 21 * 60 * 60, '라급': 34 * 60 * 60 + 30 * 60,
                   '마급': 43 * 60 * 60 + 30 * 60, '바급': 45 * 60 * 60}
        alpha_B = {'2급': 3 * 60 * 60, '3급': 15 * 60 * 60, '4급': 33 * 60 * 60, '5급': 48 * 60 * 60,
                   '6급': 52 * 60 * 60 + 30 * 60,
                   '가급': 12 * 60 * 60, '나급': 21 * 60 * 60, '다급': 34 * 60 * 60 + 30 * 60, '라급': 43 * 60 * 60 + 50 * 60,
                   '마급': 52 * 60 * 60 + 30 * 60, '바급': 63 * 60 * 60}
        if group == 'A':
            if total_secs <= alpha_A[grade]:
                money_secs = total_secs
                vacation_secs = 0
            else:
                money_secs = alpha_A[grade]
                vacation_secs = total_secs - alpha_A[grade]
        else:
            if total_secs <= alpha_B[grade]:
                money_secs = total_secs
                vacation_secs = 0
            else:
                money_secs = alpha_B[grade]
                vacation_secs = total_secs - alpha_B[grade]
        return self.secs_to_hms(vacation_secs)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "TimeMachine"))
        self.btn_upload.setText(_translate("MainWindow", "Data Upload"))
        self.lbl_start.setText(_translate("MainWindow", "Select month"))
        self.lbl_end.setText(_translate("MainWindow", "Select Name"))
        self.btn_save_excel.setText(_translate("MainWindow", "Save Excel"))
        self.btn_path.setText(_translate("MainWindow", "PATH"))
        self.btn_preview.setText(_translate("MainWindow", "Preview"))
        self.btn_postview.setText(_translate("MainWindow", "Postview"))

class MyTableModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        #            return self.header[col]
        return None


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

