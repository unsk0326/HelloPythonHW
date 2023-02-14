# -*- coding: utf-8 -*-

import pywinauto as pwa
import re
import pygame
import pathlib
import os
import pprint
import time
from ctypes import windll
from slacker import Slacker
import cv2
import cv2.cv2 as cv2
import keyboard as keys
import mouse as mo
import numpy
import numpy as np
import pyautogui as pa
import pywinauto as pwa
import win32clipboard
from PIL import ImageGrab
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
import json

slack_token = 'xoxb-902393044309-4362724146146-EyPY8UogBG5PuHVS1i0H7WHa'

user32 = windll.user32
user32.SetProcessDPIAware()

# Make Active torch window
# loc_app = [440,105]
# loc_left_return  = [349,639]
# loc_try  = [918,657]
# loc_size_lt_win = [567,180]
# loc_size_rb_win = [862,546]
# loc_craft_win_lt_win = [816,632]
# loc_craft_win_rb_win = [1020, 700]
loc_app = [152,39]
loc_left_return  = [526,819]
loc_try  = [1412,858]
loc_size_lt_win = [894,160]
loc_size_rb_win = [1341,687]
loc_craft_win_lt_win = [1302,824]
loc_craft_win_rb_win = [1595, 902]
loc_map_win_lt_win = [235,299]
loc_map_win_rb_win = [254,325]
loc_map_reroll = [264,720]
def setActiveTorch():
    app = pwa.application.Application()
    # app.connect(title_re='Path of Exile')
    app.connect(title_re='Torchlight')
    app_dialog = app.window()
    app_dialog.set_focus()

def get_linenumber():
    cf = currentframe()
    return cf.f_back.f_lineno

class Crafter():

    def __init__(self):

        self.slack = WebClient(token=slack_token)
        import configparser
        from ast import literal_eval
        self.run = False
        SECTION_HRALDRY = 'heraldry'
        configFile = os.path.dirname(os.path.realpath(__file__)) + '\\' + 'config.cfg'
        print("file:" + configFile)
        config = configparser.ConfigParser()
        config.read(configFile)

        self.minus_opt_picture = config['torch']['minus_opt_picture']
        self.plus_opt_picture = config['torch']['plus_opt_picture']
        self.gray_opt_picture = config['torch']['gray_opt_picture']
        self.up_gray_opt_picture = config['torch']['up_gray_opt_picture']
        self.blue_opt_picture = config['torch']['blue_opt_picture']
        self.purple_opt_picture = config['torch']['purple_opt_picture']
        self.less_material_picture = config['torch']['less_material_picture']
        self.t1_opt_picture = config['torch']['T1_opt_picture']
        self.wjsdml_opt_picture = config['torch']['wjsdml_opt_picture']
        self.fireskill_opt_picture = config['torch']['fireskill_opt_picture']
        self.num8_opt_picture = config['torch']['number8_opt_picture']
        self.confidence_alters =  0.87

        pygame.mixer.init()
        BUFFER = 3072  # audio buffer size, number of samples since pygame 1.8.
        freq, size, chan = pygame.mixer.get_init()
        pygame.mixer.init(freq, size, chan, BUFFER)
    def grabImage(self, checking_region):
        boxRegions = []
        x, y, w, h = checking_region
        mon = {'top': y, 'left': x, 'width': w, 'height': h}
        imgGrab = ImageGrab.grab(bbox=(x, y, x+w, y+h))
        return imgGrab
    def findImage(self, imgGrab, checking_region, templateName, show=0, confidence=0.6):
        boxRegions = []
        x, y, w, h = checking_region
        mon = {'top': y, 'left': x, 'width': w, 'height': h}
        # sct = mss.mss()
        # sct.grab(mon)
        # img = Image.frombytes('RGB', (w, h), sct.grab(mon).rgb)
        # imgGrab = ImageGrab.grab(bbox=(x, y, x+w, y+h))
        img = cv2.cvtColor(numpy.array(imgGrab), cv2.COLOR_RGB2BGR)
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        #gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        colored = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)

        imgPath = os.path.dirname(os.path.realpath(__file__)) + '\\' + templateName
        template = cv2.imread(imgPath)

        w, h = template.shape[:-1]

        #res = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
        res = cv2.matchTemplate(colored, template, cv2.TM_CCOEFF_NORMED)
        threshold = confidence
        loc = np.where(res >= threshold)

        for pt in zip(*loc[::-1]):
            if show == 1:
                cv2.rectangle(frame, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)
            # realX = self.inventorySize[0] + pt[0]
            # realY = self.inventorySize[1] + pt[1]
            realX = checking_region[0] + pt[0]
            realY = checking_region[1] + pt[1]
            boxRegions.append((realX, realY, w, h))

        if show == 1:
            cv2.imshow('image', frame)
            cv2.waitKey(0)
            cv2.destroyAllWindows()

        return boxRegions
    def check_pic(self,gray, capture_loc,pic):
        ret = self.findImage(gray,capture_loc, pic, show=0,
                                        confidence=self.confidence_alters)
        k = len(ret)
        if (k == 0):
            return False
        else:
            # print("confi and ret=", self.confidence_alters, k)

            # for region in ret:
            #     print("alter x y =", region[0], region[1])
            # print("previous: ", self.loc_orb_alter)
            self.loc_orb_alter = (ret[0][0], ret[0][1])
            # print("after: ", self.loc_orb_alter)
            return True
    def play_mp3(self):
        mp3_loc = str(pathlib.Path().absolute()) + "\\very.mp3"
        print(mp3_loc)
        pygame.init()
        pygame.mixer.init()
        clock = pygame.time.Clock()
        pygame.mixer.music.load(mp3_loc)
        pygame.mixer.music.set_volume(0.02)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            clock.tick(500)
    def hit_item(self, targetPoint):
        mo.move(targetPoint[0], targetPoint[1])
        time.sleep(0.05)
        mo.click()
        time.sleep(0.06)

    def pureimg(self,data1):
        data1 = '[{"text": "", "image_url": "' + data1 + '"}]'
        data1 = [json.loads(data1[1:-1])]
        return data1

    def create_T1_option(self, check_plus):
        keyState = 0
        self.run = False

        it = Crafter()

        while True:
            time.sleep(0.05)
            stop_key = keys.is_pressed('F1')
            start_key = keys.is_pressed('F2')
            exit = keys.is_pressed('F3')
            if (exit == True):
                break

            if start_key == True:
                from win32gui import FindWindow, GetWindowRect

                window_handle = FindWindow(None, "Torchlight: Infinite  ")
                window_rect = GetWindowRect(window_handle)

                loc_app[0] = window_rect[0]
                loc_app[1] = window_rect[1]

                print(" Let's create rare~")
                print(" Let's create rare~")
                print(" Let's create rare~")
                print("app position = {}, {}".format(loc_app[0], loc_app[1]))
                self.loc_capture_size = [loc_app[0]+loc_size_lt_win[0],loc_app[1]+loc_size_lt_win[1],loc_size_rb_win[0]-loc_size_lt_win[0],loc_size_rb_win[1]-loc_size_lt_win[1] ]
                print(self.loc_capture_size)
                self.loc_capture_craft_size = [loc_app[0] + loc_craft_win_lt_win[0], loc_app[1] + loc_craft_win_lt_win[1],
                                         loc_craft_win_rb_win[0] - loc_craft_win_lt_win[0],
                                         loc_craft_win_rb_win[1] - loc_craft_win_lt_win[1]]
                self.try_button = [loc_app[0] + loc_try[0], loc_app[1] + loc_try[1]]
                self.left_button = [loc_app[0] + loc_left_return[0], loc_app[1] + loc_left_return[1]]

                self.run = True
            if stop_key == True:
                self.run = False

            if self.run == True:
                time.sleep(0.5)
                #self.hit_item()
                #mo.move(loc_app[0]+loc_try[0],loc_app[1]+loc_try[1])
                should_click_left = False
                img = self.grabImage(self.loc_capture_size)
                # check +1 not exist
                # yes -> click left
                if check_plus:
                    ret = self.check_pic(img, self.loc_capture_size, self.plus_opt_picture)
                    if(ret == False):
                        print("! no plus option issue")
                        should_click_left = True

                # check2 minus exist
                # yes -> click left
                if should_click_left == False:
                    ret = self.check_pic(img, self.loc_capture_size, self.minus_opt_picture)
                    if (ret == True):
                        print("## minus option issue")
                        should_click_left = True
                    # check3 gray exist
                    # yes -> click left
                if should_click_left == False:
                    ret = self.check_pic(img, self.loc_capture_size, self.gray_opt_picture)
                    if (ret == True):
                        print("1 gray issue")
                        should_click_left = True
                if should_click_left == False:
                    ret = self.check_pic(img, self.loc_capture_size, self.up_gray_opt_picture)
                    if (ret == True):
                        print("11 upper gray option issue")
                        should_click_left = True

                # check4 blue
                # yes -> click left
                if should_click_left == False:
                    ret = self.check_pic(img, self.loc_capture_size, self.blue_opt_picture)
                    if (ret == True):
                        print("2 blue option issue")
                        should_click_left = True
                #check5 purple
                # yes -> click left
                if should_click_left == False:
                    ret = self.check_pic(img, self.loc_capture_size, self.purple_opt_picture)
                    if (ret == True):
                        print("3 purple option issue")
                        should_click_left = True

                if should_click_left:
                    self.hit_item(self.left_button)
                    time.sleep(0.1)
                    img = self.grabImage(self.loc_capture_craft_size)
                    ret = self.check_pic(img, self.loc_capture_craft_size, self.less_material_picture)
                    if ret:
                        print("no material issue")
                        self.slack.chat_postMessage(channel='general',
                                                    text='less material check it please')

                        self.run = False
                    else:

                        self.hit_item(self.try_button)
                else:
                    print("hello it's done")
                    self.play_mp3()

                    #self.slack.chat_postMessage(channel='general',text='finished check result please')
                    img2 = cv2.cvtColor(numpy.array(img), cv2.COLOR_RGB2BGR)
                    payoff = './img.png'
                    cv2.imwrite(payoff, img2)
                    print(payoff)
                    # It gives cross OS compatibility on filepath.
                    response = self.slack.files_upload(
                        file='./img.png',
                        initial_comment='finished check result',
                        channels='general'
                    )
                    self.run = False
    def reroll_T1_option(self, wjsdml=False):
        keyState = 0
        self.run = False

        it = Crafter()

        while True:
            time.sleep(0.05)
            stop_key = keys.is_pressed('F1')
            start_key = keys.is_pressed('F2')
            exit = keys.is_pressed('F3')

            if (exit == True):
                break

            if start_key == True:
                from win32gui import FindWindow, GetWindowRect

                window_handle = FindWindow(None, "Torchlight: Infinite  ")
                window_rect = GetWindowRect(window_handle)

                loc_app[0] = window_rect[0]
                loc_app[1] = window_rect[1]

                print(" Let's create rare~")
                print(" Let's create rare~")
                print(" Let's create rare~")
                print("app position = {}, {}".format(loc_app[0], loc_app[1]))
                self.loc_capture_size = [loc_app[0]+loc_size_lt_win[0],loc_app[1]+loc_size_lt_win[1],loc_size_rb_win[0]-loc_size_lt_win[0],loc_size_rb_win[1]-loc_size_lt_win[1] ]
                print(self.loc_capture_size)
                self.loc_capture_craft_size = [loc_app[0] + loc_craft_win_lt_win[0], loc_app[1] + loc_craft_win_lt_win[1],
                                         loc_craft_win_rb_win[0] - loc_craft_win_lt_win[0],
                                         loc_craft_win_rb_win[1] - loc_craft_win_lt_win[1]]
                self.try_button = [loc_app[0] + loc_try[0], loc_app[1] + loc_try[1]]
                self.left_button = [loc_app[0] + loc_left_return[0], loc_app[1] + loc_left_return[1]]

                self.run = True
            if stop_key == True:
                self.run = False

            if self.run == True:
                time.sleep(0.45)
                #self.hit_item()
                #mo.move(loc_app[0]+loc_try[0],loc_app[1]+loc_try[1])
                should_click_left = False
                img = self.grabImage(self.loc_capture_size)

                # check t1
                # no -> click left
                ret = self.check_pic(img, self.loc_capture_size, self.t1_opt_picture)
                ret2 = self.check_pic(img, self.loc_capture_size, self.wjsdml_opt_picture)
                if (ret != True):
                    print("1 t1 option is not found")
                    should_click_left = True
                if wjsdml == True and ret2 == False:
                    print("1 전의 option is not found")
                    should_click_left = True

                if should_click_left:
                    self.hit_item(self.left_button)
                    time.sleep(0.1)
                    img = self.grabImage(self.loc_capture_craft_size)
                    ret = self.check_pic(img, self.loc_capture_craft_size, self.less_material_picture)
                    if ret:
                        print("no material issue")
                        self.slack.chat_postMessage(channel='general',
                                                    text='less material check it please')

                        self.run = False
                    else:

                        self.hit_item(self.try_button)
                else:
                    print("hello it's done")
                    self.play_mp3()
                    if False:


                        #self.slack.chat_postMessage(channel='general',text='finished check result please')
                        img2 = cv2.cvtColor(numpy.array(img), cv2.COLOR_RGB2BGR)
                        payoff = './img.png'
                        cv2.imwrite(payoff, img2)
                        print(payoff)
                        # It gives cross OS compatibility on filepath.
                        response = self.slack.files_upload(
                            file='./img.png',
                            initial_comment='finished check result',
                            channels='general'
                        )
                    self.run = False
    def reroll_T1_option2(self, fireskill=False):
        keyState = 0
        self.run = False

        it = Crafter()

        while True:
            time.sleep(0.05)
            stop_key = keys.is_pressed('F1')
            start_key = keys.is_pressed('F2')
            exit = keys.is_pressed('F3')

            if (exit == True):
                break

            if start_key == True:
                from win32gui import FindWindow, GetWindowRect

                window_handle = FindWindow(None, "Torchlight: Infinite  ")
                window_rect = GetWindowRect(window_handle)

                loc_app[0] = window_rect[0]
                loc_app[1] = window_rect[1]

                print(" Let's create rare~")
                print(" Let's create rare~")
                print(" Let's create rare~")
                print("app position = {}, {}".format(loc_app[0], loc_app[1]))
                self.loc_capture_size = [loc_app[0]+loc_size_lt_win[0],loc_app[1]+loc_size_lt_win[1],loc_size_rb_win[0]-loc_size_lt_win[0],loc_size_rb_win[1]-loc_size_lt_win[1] ]
                print(self.loc_capture_size)
                self.loc_capture_craft_size = [loc_app[0] + loc_craft_win_lt_win[0], loc_app[1] + loc_craft_win_lt_win[1],
                                         loc_craft_win_rb_win[0] - loc_craft_win_lt_win[0],
                                         loc_craft_win_rb_win[1] - loc_craft_win_lt_win[1]]
                self.try_button = [loc_app[0] + loc_try[0], loc_app[1] + loc_try[1]]
                self.left_button = [loc_app[0] + loc_left_return[0], loc_app[1] + loc_left_return[1]]

                self.run = True
            if stop_key == True:
                self.run = False

            if self.run == True:
                time.sleep(0.48)
                #self.hit_item()
                #mo.move(loc_app[0]+loc_try[0],loc_app[1]+loc_try[1])
                should_click_left = False
                img = self.grabImage(self.loc_capture_size)

                # check t1
                # no -> click left
                ret = self.check_pic(img, self.loc_capture_size, self.t1_opt_picture)
                ret2 = self.check_pic(img, self.loc_capture_size, self.fireskill_opt_picture)
                if (ret != True):
                    print("1 t1 option is not found")
                    should_click_left = True
                if fireskill == True and ret2 == False:
                    print("1 화염 option is not found")
                    should_click_left = True

                if should_click_left:
                    self.hit_item(self.left_button)
                    time.sleep(0.3)
                    img = self.grabImage(self.loc_capture_craft_size)
                    ret = self.check_pic(img, self.loc_capture_craft_size, self.less_material_picture)
                    if ret:
                        print("no material issue")
                        # self.slack.chat_postMessage(channel='general',
                        #                             text='less material check it please')

                        self.run = False
                    else:

                        self.hit_item(self.try_button)
                else:
                    print("hello it's done")
                    self.play_mp3()
                    if False:

                        pass
                        #self.slack.chat_postMessage(channel='general',text='finished check result please')
                        # img2 = cv2.cvtColor(numpy.array(img), cv2.COLOR_RGB2BGR)
                        # payoff = './img.png'
                        #cv2.imwrite(payoff, img2)
                        # It gives cross OS compatibility on filepath.
                        # response = self.slack.files_upload(
                        #     file='./img.png',
                        #     initial_comment='finished check result',
                        #     channels='general'
                        # )
                    self.run = False
    def reroll_map_quanty(self, fireskill=False):
        keyState = 0
        self.run = False

        it = Crafter()

        while True:
            time.sleep(0.05)
            stop_key = keys.is_pressed('F1')
            start_key = keys.is_pressed('F2')
            exit = keys.is_pressed('F3')

            if (exit == True):
                break

            if start_key == True:
                from win32gui import FindWindow, GetWindowRect

                window_handle = FindWindow(None, "Torchlight: Infinite  ")
                window_rect = GetWindowRect(window_handle)

                loc_app[0] = window_rect[0]
                loc_app[1] = window_rect[1]

                print(" Let's create rare~")
                print(" Let's create rare~")
                print(" Let's create rare~")
                print("app position = {}, {}".format(loc_app[0], loc_app[1]))
                self.loc_map_quanty_size = [loc_app[0]+loc_map_win_lt_win[0],loc_app[1]+loc_map_win_lt_win[1],loc_map_win_rb_win[0]-loc_map_win_lt_win[0],loc_map_win_rb_win[1]-loc_map_win_lt_win[1] ]

                self.try_button = [loc_app[0] + loc_map_reroll[0], loc_app[1] + loc_map_reroll[1]]


                self.run = True
            if stop_key == True:
                self.run = False

            if self.run == True:
                time.sleep(0.1)
                #self.hit_item()
                #mo.move(loc_app[0]+loc_try[0],loc_app[1]+loc_try[1])
                should_reroll = False
                img = self.grabImage(self.loc_map_quanty_size)

                # check t1
                # no -> click left
                ret = self.check_pic(img, self.loc_map_quanty_size, self.num8_opt_picture)

                if (ret != True):
                    print("1 num8 option is not found")
                    should_reroll = True
                else:
                    print("quant is enough")
                    self.run = False

                if should_reroll:
                    self.hit_item(self.try_button)


if __name__ == "__main__":
    run = 6
    cft = Crafter()
    if run == 1:
        setActiveTorch()
        print("torch is launched")
        cft.create_T1_option(True)
    elif run == 2:
        setActiveTorch()
        print("torch is launched")
        cft.create_T1_option(False)
    elif run == 3:
        setActiveTorch()
        print("torch is launched")
        cft.reroll_T1_option()
    elif run == 4:
        setActiveTorch()
        print("torch is launched")
        cft.reroll_T1_option(True)
    elif run == 5:
        setActiveTorch()
        print("torch is launched")
        cft.reroll_T1_option2(True)
    elif run == 6:
        setActiveTorch()
        print("torch is launched, quanty")
        cft.reroll_map_quanty()
