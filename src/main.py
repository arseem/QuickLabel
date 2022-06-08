from pandas.io import excel
import qrcode
import barcode
from barcode.writer import ImageWriter
import PIL.Image
import PIL.ImageFont
import PIL.ImageDraw
import ctypes
import os
import sys
import pandas as pd
import matplotlib.font_manager as fm
from tkinter import filedialog
from tkinter import Tk
from win32api import GetSystemMetrics

import kivy
from kivy.app import App
from kivy.uix.widget import Widget
from kivy.uix.boxlayout import BoxLayout
from kivy.lang import Builder
from kivy.properties import ObjectProperty
from kivy.core.window import Window
from kivy.core.image import Image as CoreImage
from kivy.clock import Clock, mainthread
from kivy.uix.popup import Popup
from kivy.uix.button import Button
from kivy.uix.behaviors import ToggleButtonBehavior
from kivy.uix.image import Image
from io import BytesIO
import threading
import time
from functools import partial
import win32print
import win32api

ctypes.windll.shcore.SetProcessDpiAwareness(1)


Builder.load_string('''
<MyGrid>
    kod:kod
    opis:opis
    serial:serial

    GridLayout:
        cols:3
        size: root.width, root.height

        FloatLayout:
            #orientation: 'vertical'

            Image:
                pos:0,0
                source:""
                id: etykieta
                allow_stretch: True
                keep_ratio: True
                texture:root.tex

            Button:
                id:left_but
                pos_hint:{'x':0, 'center_y':.5}
                size_hint:.03,.2
                text: '<'
                opacity:0
                on_release: root.go_left()
                font_size: 0.1*self.height
                disabled:True
            
            Button:
                id:right_but
                pos_hint:{'right':1, 'center_y':.5}
                size_hint:.03,.2
                text: '>'
                opacity:0
                on_release: root.go_right()
                font_size: 0.1*self.height
                disabled:True

            
            Button:
                id:last_one
                pos_hint:{'center_x':.5, 'bottom_y':0}
                size_hint:.4,.04
                text: 'COFNIJ OSTATNI'
                opacity:0.8
                on_release: root.delete_one()
                font_size: 0.45*self.height
                disabled:True



        
        GridLayout:
            cols:2

            BoxLayout:
                id:iface
                orientation:'vertical'
                spacing:self.height * 0.02
                padding:(self.width * 0.11,self.width * 0.11,0,self.width * 0.11)

                Label:
                    id: kodlabel
                    text: 'NUMER SERYJNY'
                    size_hint:(.2,.2)
                    pos_hint:{'center_x':.5}
                    font_size:0.9*self.height
                
                TextInput:
                    id: kod
                    multiline: False
                    size_hint:(.9,.4)
                    pos_hint:{'center_x':.5}
                    font_size: 0.95*kodlabel.height
                    write_tab: False
                    on_text_validate: root.on_enter()
                    disabled: True if reader_toggle.state=='down' and lock_numer.state=='down' else False


                Label:
                    text: 'OPIS'
                    size_hint:(.2,.2)
                    pos_hint:{'center_x':.5}
                    font_size:0.9*self.height

                TextInput:
                    id: opis
                    multiline: False
                    size_hint:(.9,.4)
                    pos_hint:{'center_x':.5}
                    font_size: 0.95*kodlabel.height
                    write_tab: False
                    on_text_validate: root.on_enter()
                    disabled: True if reader_toggle.state=='down' and lock_opis.state=='down' else False


                Label:
                    text: 'SKŁAD'    
                    size_hint:(.2,.2)
                    pos_hint:{'center_x':.5} 
                    font_size:0.9*self.height
       
                    
                TextInput:
                    id: serial
                    multiline: False
                    size_hint:(.9,.4)
                    pos_hint:{'center_x':.5}
                    font_size: 0.95*kodlabel.height
                    write_tab: False
                    on_text_validate: root.on_enter()
                    disabled: True if reader_toggle.state=='down' and lock_sklad.state=='down' else False
            

                BoxLayout:
                    size_hint:1,.2


                GridLayout:
                    cols:3
                    size_hint:(1,.4)

                    Button:
                        text:'WYCZYŚĆ'
                        size_hint:(.4,.5)
                        text_size: self.size
                        halign:'center'
                        valign: 'center'
                        font_size:0.3*self.height
                        on_release: root.clear()
                    
                    Button:
                        text:'DODAJ'
                        size_hint:(.4,.5)
                        text_size: self.size
                        halign:'center'
                        valign: 'center'
                        font_size:0.3*self.height
                        on_release: root.press_add()

                    Button:
                        id:but
                        text:'WCZYTAJ Z ARKUSZA'
                        size_hint:(.4,.5)
                        text_size: self.size
                        halign:'center'
                        valign: 'center'
                        font_size:0.3*self.height
                        on_release: root.thread_excel()
                

                BoxLayout:
                    size_hint:1,.2


                GridLayout:
                    cols:3

                    Button:
                        id:save_pdf
                        text:'ZAPISZ PDF'
                        size_hint:(None,None)
                        size:(but.width, but.height) if not print_but.is_open else (0, 0)
                        opacity: 1 if not print_but.is_open else 0
                        text_size: self.size
                        halign:'center'
                        valign: 'center'
                        pos_hint:{'center_x':.5} 
                        font_size:0.3*self.height
                        on_release: root.save_pdf()
                        disabled:True
                        
                    Spinner:
                        id:print_but
                        text:'DRUKUJ PDF'
                        size_hint:(None,None)
                        size:(but.width*3, but.height) if self.is_open else (but.width, but.height)
                        text_size: self.size
                        pos_hint:{'center_x':.5}
                        font_size:0.3*self.height
                        halign:'center'
                        valign: 'center' 
                        on_release: root.print_pdf()
                        disabled:True
                        on_text:root.printer_spinner_clicked(print_but.text)
                            
                    Button:
                        id:save_rap
                        text:'ZAPISZ RAPORT'
                        size_hint:(None,None)
                        size:(but.width, but.height) if not print_but.is_open else (0, 0)
                        opacity: 1 if not print_but.is_open else 0
                        text_size: self.size
                        halign:'center'
                        valign: 'center'
                        pos_hint:{'center_x':.5}
                        font_size:0.3*self.height
                        on_release: root.save_excel()
                        disabled:True
                            

                BoxLayout:
                    orientation:'horizontal'

                BoxLayout:
                    orientation:'horizontal'
                    size_hint: 1, .5

                BoxLayout:
                    orientation:'horizontal'
                    size_hint: 1, .5

                    ToggleButton:
                        id:reader_toggle
                        text:'TRYB CZYTNIKA'
                        pos_hint:{'center_x':.5}
                        font_size:0.3*self.height
                        text_size: self.size
                        halign:'center'
                        valign: 'center'

            GridLayout:
                cols:1
                size_hint:(.1,1)
                
                BoxLayout:
                    size_hint:1,.13

                AnchorLayout:
                    size_hint:1,.18
                    ToggleButton:
                        id:lock_numer
                        border: 0, 0, 0, 0
                        size_hint: None, None
                        size:(but.height, but.height*5//6) if self.state=='normal' else (but.height*2//3, but.height*5//6)
                        background_normal: root.path_res + '/normal.png'
                        background_down: root.path_res + '/down.png'
                        opacity: 0 if reader_toggle.state=='normal' else 1
                        disabled:True if reader_toggle.state=='normal' else False
                
                AnchorLayout:
                    size_hint:1,.18
                    ToggleButton:
                        id:lock_opis
                        border: 0, 0, 0, 0
                        size_hint: None, None
                        size:(but.height, but.height*5//6) if self.state=='normal' else (but.height*2//3, but.height*5//6)
                        background_normal: root.path_res + '/normal.png'
                        background_down: root.path_res + '/down.png'
                        opacity: 0 if reader_toggle.state=='normal' else 1
                        disabled:True if reader_toggle.state=='normal' else False 

                AnchorLayout:
                    size_hint:1,.18
                    ToggleButton:
                        id:lock_sklad
                        border: 0, 0, 0, 0
                        size_hint: None, None
                        size:(but.height, but.height*5//6) if self.state=='normal' else (but.height*2//3, but.height*5//6)
                        background_normal: root.path_res + '/normal.png'
                        background_down: root.path_res + '/down.png'
                        opacity: 0 if reader_toggle.state=='normal' else 1
                        disabled:True if reader_toggle.state=='normal' else False
                
                BoxLayout:


<MyPopup@Popup>:
    orientation:'vertical'
    BoxLayout:
        size_hint:.5,.5
    Label:
        text: 'CZEKAJ'    
        size_hint:(.4,.4)
        pos_hint:{'center_x':.5, 'center_y':.5} 
        font_size:0.9*self.height

    ProgressBar:
        id:loading
        color:(0,0,0,1)
        min:0
        max:100
        val:0
        pos_hint:{'center_x':.5, 'center_y':.5}
        size:0,0           


<MyError@Popup>:
    orientation:'vertical'
    auto_dismiss: True
    BoxLayout:
        size_hint:.5,.5
    Label:
        text: 'ZAPEŁNIJ WSZYSTKIE POLA'    
        size_hint:(.3,.3)
        pos_hint:{'center_x':.5, 'center_y':.5} 
        halign: 'center'
        valign: 'center'
        font_size:0.9*self.height

    BoxLayout:
        size_hint:.5,.5

''')
Window.size = (GetSystemMetrics(0)//1.5, GetSystemMetrics(0)//1.5*297//420)
def_printer = win32print.GetDefaultPrinter()


class MyPopup(BoxLayout):
    pass


class MyError(BoxLayout):
    pass


class ImageButton(Image, ToggleButtonBehavior):
    pass


class MyGrid(Widget):

    bl = PIL.Image.new('RGB', (2480//2, 472), color = 'black')
    dt = BytesIO()
    bl.save(dt, format='png')
    dt.seek(0)
    black = CoreImage(BytesIO(dt.read()), ext='png').texture

    A4w = 2480
    A4h = 3508
    hgrid = 501
    SEPARATOR = ' '
    IMAGESLIST = []
    DISPLAYIMAGESLIST = []
    DISPLAYINDEX = -1
    ALLDATA = {'NUMERY':[], 'OPISY':[], 'SKLADY':[]}
    CURRENT_ON_GRID = (0,0)
    LAST_GRID = (0,0)
    CURRENT_PAGE = False
    ORDER = [(0,0), (0,1), (0,2), (0,3), (0,4), (0,5), (0,6), (1,0), (1,1), (1,2), (1,3), (1,4), (1,5), (1,6)]

    kod = ObjectProperty(None)
    opis = ObjectProperty(None)
    serial = ObjectProperty(None)
    tex = black

    path_res = './res'

    show = MyPopup()
    popupWindow = Popup(title='', content=show, size_hint=(.5, .2))

    error = MyError()
    error_window = Popup(title='', content=error, size_hint=(.5, .2))


    def press_add(self):
        if self.kod.text and self.opis.text and self.serial.text:
            self.start(self.kod.text, self.opis.text, self.serial.text)
            self.display()

        else:
            self.error_window.open()

    
    def on_enter(self):
        if not self.kod.text:
            if self.ids.lock_numer.state != 'down':
                self.ids.kod.focus = True
            
            else:
                self.error_window.open()
        
        elif not self.opis.text:
            if not self.ids.lock_opis.state == 'down':
                self.ids.opis.focus = True
            
            else:
                self.error_window.open()
        
        elif not self.serial.text:
            if not self.ids.lock_sklad.state == 'down':
                self.ids.serial.focus = True
            
            else:
                self.error_window.open()

        else:            
            self.press_add()

            focused = False
            if not self.ids.lock_numer.state == 'down':
                self.ids.kod.text = ''
                Clock.schedule_once(lambda *args: self.set_focus(self.ids.kod), .5)
                focused = True
            
            if not self.ids.lock_opis.state == 'down':
                self.ids.opis.text = ''
                if not focused:
                    Clock.schedule_once(lambda *args: self.set_focus(self.ids.opis), .5)
                    focused = True
            
            if not self.ids.lock_sklad.state == 'down':
                self.ids.serial.text = ''
                if not focused:
                    Clock.schedule_once(lambda *args: self.set_focus(self.ids.serial), .5)
                    focused = True

    
    def set_focus(self, object):
        object.focus = True


    @mainthread
    def display(self):
        if self.DISPLAYIMAGESLIST:
            self.ids["etykieta"].texture = self.DISPLAYIMAGESLIST[self.DISPLAYINDEX]
            self.ids["etykieta"].reload()
            self.ids["save_pdf"].disabled = False
            self.ids["print_but"].disabled = False
            self.ids["save_rap"].disabled = False
            self.ids["last_one"].disabled = False
        
        else:
            self.ids["etykieta"].texture = self.black
            self.ids["etykieta"].reload()
            self.ids["save_pdf"].disabled = True
            self.ids["print_but"].disabled = True
            self.ids["save_rap"].disabled = True
            self.ids["last_one"].disabled = True

        
        if len(self.DISPLAYIMAGESLIST)>1:
            self.ids["left_but"].disabled = False if self.DISPLAYINDEX!=0 else True  
            self.ids["right_but"].disabled = False if self.DISPLAYINDEX!=len(self.DISPLAYIMAGESLIST)-1 and self.DISPLAYINDEX!=-1 else True
            self.ids["left_but"].opacity = .8 if self.DISPLAYINDEX!=0 else 0  
            self.ids["right_but"].opacity = .8 if self.DISPLAYINDEX!=len(self.DISPLAYIMAGESLIST)-1 and self.DISPLAYINDEX!=-1 else 0
        
        else:
            self.ids["left_but"].disabled = True
            self.ids["right_but"].disabled = True
            self.ids["left_but"].opacity = 0
            self.ids["right_but"].opacity = 0
        
                


    
    def start(self, numerText, opisText, skladText, excel=False):

        numer = numerText
        opis = opisText
        sklad = skladText


        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=1000//33,
            border = 1000//122
        )

        qr.clear()
        qr.add_data(f'{numer}{self.SEPARATOR}{opis}{self.SEPARATOR}{sklad}')
        qr.make(fit=True)

        qr_img_org = qr.make_image(fill_color="black", back_color="white").convert('RGB')
        qr_img = qr_img_org.resize((self.hgrid, self.hgrid), PIL.Image.ANTIALIAS)

        illegals = 'ąęćłóżźśńŃĄĘĆŁÓŻŹŚ'
        to_legals = 'aeclozzsnNAECLOZZS'

        numer_clear, opis_clear, sklad_clear = str(numer), str(opis), str(sklad)
        for n in range(len(illegals)):
            numer_clear = numer_clear.replace(illegals[n], to_legals[n])
            opis_clear = opis_clear.replace(illegals[n], to_legals[n])
            sklad_clear = sklad_clear.replace(illegals[n], to_legals[n])

        bar = barcode.get('CODE128', f'{numer_clear}', writer = ImageWriter())
        bar_img_org = bar.render(writer_options = {'font_path':fm.findfont(fm.FontProperties(family='arial'))})
        k = bar_img_org.height/bar_img_org.width
        bar_img1 = bar_img_org.resize((2480//2-500, 472//5), PIL.Image.ANTIALIAS)
        bar_img1 = bar_img1.crop((0,0,2480//2-500,472//6))

        bar = barcode.get('CODE128', f'{opis_clear}', writer = ImageWriter())
        bar_img_org = bar.render(writer_options = {'font_path':fm.findfont(fm.FontProperties(family='arial'))})
        k = bar_img_org.height/bar_img_org.width
        bar_img2 = bar_img_org.resize((2480//2-500, 472//5), PIL.Image.ANTIALIAS)
        bar_img2 = bar_img2.crop((0,0,2480//2-500,472//6))

        bar = barcode.get('CODE128', f'{sklad_clear}', writer = ImageWriter())
        bar_img_org = bar.render(writer_options = {'font_path':fm.findfont(fm.FontProperties(family='arial'))})
        k = bar_img_org.height/bar_img_org.width
        bar_img3 = bar_img_org.resize((2480//2-500, 472//5), PIL.Image.ANTIALIAS)
        bar_img3 = bar_img3.crop((0,0,2480//2-500,472//6))


        #fin_display = fin.resize((fin.width, fin.height), PIL.Image.ANTIALIAS)
        self.appender(qr_img, bar_img1, bar_img2, bar_img3, numer, opis, sklad)


    @mainthread
    def appender(self, qr_img, bar_img1, bar_img2, bar_img3, numer, opis, sklad):
        self.ALLDATA['NUMERY'].append(numer)
        self.ALLDATA['OPISY'].append(opis)
        self.ALLDATA['SKLADY'].append(sklad)
        
        fin = PIL.Image.new('RGB', (self.A4w, self.A4h), color='white') if not self.CURRENT_PAGE else self.CURRENT_PAGE
        blank = PIL.Image.new('RGB', (2480//2, 472), color = 'white')
        draw = PIL.ImageDraw.Draw(blank)

        draw.text((472, self.hgrid//20+70), f'{numer}', (0,0,0), font=PIL.ImageFont.truetype(fm.findfont(fm.FontProperties(family='arial', weight = 'bold')), 55))
        draw.text((472, 7*self.hgrid//20+70), f'{opis}', (0,0,0), font=PIL.ImageFont.truetype(fm.findfont(fm.FontProperties(family='arial')), 55))
        draw.text((472, 13*self.hgrid//20+70), f'{sklad}', (0,0,0), font=PIL.ImageFont.truetype(fm.findfont(fm.FontProperties(family='arial')), 40))

        n = self.CURRENT_ON_GRID[0]
        m = self.CURRENT_ON_GRID[1]
        self.LAST_GRID = (n, m)
        fin.paste(blank, (472//10+n*self.A4w//2, m*self.hgrid))
        fin.paste(qr_img, (0+n*self.A4w//2,0+m*self.hgrid))
        fin.paste(bar_img1, (472+n*self.A4w//2, self.hgrid//20 + m*self.hgrid))
        fin.paste(bar_img2, (472+n*self.A4w//2, 7*self.hgrid//20 + m*self.hgrid))
        fin.paste(bar_img3, (472+n*self.A4w//2, 13*self.hgrid//20 + m*self.hgrid))

        self.CURRENT_ON_GRID = (self.CURRENT_ON_GRID[0], self.CURRENT_ON_GRID[1]+1) if self.CURRENT_ON_GRID[1] < 6 else (1, 0) if self.CURRENT_ON_GRID[0] == 0 else (0, 0)
        self.CURRENT_PAGE = fin if not self.CURRENT_ON_GRID == (0, 0) else PIL.Image.new('RGB', (self.A4w, self.A4h), color='white')

        data = BytesIO()
        fin.save(data, format='png')
        data.seek(0)
        fin_display = CoreImage(BytesIO(data.read()), ext='png').texture

        if self.CURRENT_ON_GRID == (0,1):
            self.IMAGESLIST.append(fin)
            self.DISPLAYIMAGESLIST.append(fin_display)
            self.DISPLAYINDEX = len(self.DISPLAYIMAGESLIST)-1
        
        else:
            if self.IMAGESLIST:
                self.IMAGESLIST[-1] = fin
                self.DISPLAYIMAGESLIST[-1] = fin_display
            
            else:
                self.IMAGESLIST = [fin]
                self.DISPLAYIMAGESLIST = [fin_display]


    def save(self):
        root = Tk()
        root.withdraw()
        curr_directory = os.getcwd()
        name = filedialog.asksaveasfilename(initialdir = curr_directory, title = "Podaj nazwę pliku", filetypes = (("pliki png","*.png"),("all files","*.*")))+'.png'
        root.destroy()

        self.IMAGESLIST[self.DISPLAYINDEX].save(name)


    @mainthread
    def thread_excel(self):
        self.t = threading.Thread(target = self.load_excel)
        self.t.start()


    def load_excel(self):
        root = Tk()
        root.withdraw()
        curr_directory = os.getcwd()
        name = filedialog.askopenfilename(initialdir = curr_directory, title = "Znajdź plik arkusza", filetypes = (("pliki arkusza","*.xlsx"),("pliki arkusza","*.xls"),("all files","*.*")))
        root.destroy()

        excel_file = pd.read_excel(name, header=None, usecols=[0,1,2], names=['NUMERY','OPISY','SKLADY'])

        self.show.ids.loading.max = len(excel_file['NUMERY'].tolist())
        self.show.ids.loading.value=0
        self.popupWindow.open()

        for n in range(len(excel_file['NUMERY'].tolist())):
            Clock.schedule_once(self.update_progress)
            #Clock.schedule_once(partial(self.loop, n, excel_file))
            self.start(excel_file['NUMERY'].tolist()[n], excel_file['OPISY'].tolist()[n], excel_file['SKLADY'].tolist()[n], excel=True)
            #self.loop(n, excel_file)

        self.popupWindow.dismiss()
        self.display()

    
    def loop(self, n, excel_file, *args):
        self.start(excel_file['NUMERY'].tolist()[n], excel_file['OPISY'].tolist()[n], excel_file['SKLADY'].tolist()[n], excel=True)


    @mainthread
    def update_progress(self, *args):
        self.show.ids.loading.value+=2


    def save_excel(self):
        root = Tk()
        root.withdraw()
        curr_directory = os.getcwd()
        name = filedialog.asksaveasfilename(initialdir = curr_directory, title = "Podaj nazwę pliku", filetypes = (("pliki xlsx","*.xlsx"),("all files","*.*")))+'.xlsx'
        root.destroy()

        if name:
            to_file = pd.DataFrame(self.ALLDATA)
            to_file.to_excel(name, header=False, index=False)


    def save_pdf(self):
        root = Tk()
        root.withdraw()
        curr_directory = os.getcwd()
        name = filedialog.asksaveasfilename(initialdir = curr_directory, title = "Podaj nazwę pliku", filetypes = (("pliki pdf","*.pdf"),("all files","*.*")))+'.pdf'
        root.destroy()

        if name:
            self.IMAGESLIST[0].save(name, save_all=True, append_images=self.IMAGESLIST[1:])


    def print_pdf(self):
        printers = [x['pPrinterName'] for x in win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, None, 2)]
        default = win32print.GetDefaultPrinter()
        printers.remove(default)
        all_printers = [f'(DOMYŚLNA) {def_printer}']
        all_printers.extend(printers)
        self.ids['print_but'].values = all_printers


    def printer_spinner_clicked(self, value):
        try:
            os.remove(f'{os.getcwd()}/tmpQL.pdf')
    
        except:
            pass
        
        self.IMAGESLIST[0].save(f'{os.getcwd()}/tmpQL.pdf', save_all=True, append_images=self.IMAGESLIST[1:])
        os.system( f"attrib +h {os.getcwd()}/tmpQL.pdf" )
        value = def_printer if value == f'Domyślna drukarka\n({def_printer})' else value
        win32api.ShellExecute(
            0,
            "printto",
            f'{os.getcwd()}/tmpQL.pdf',
            f'"{value}"',
            ".",
            0
        )
        #os.startfile(f'{os.getcwd()}/tmpQL.pdf', 'print')
        self.ids['print_but'].text = 'DRUKUJ PDF'


    def clear(self):
        self.IMAGESLIST = []
        self.DISPLAYIMAGESLIST = []
        self.DISPLAYINDEX = -1
        self.CURRENT_ON_GRID = (0,0)
        self.LAST_GRID = (0,0)
        self.CURRENT_PAGE = False
        self.ALLDATA = {'NUMERY':[], 'OPISY':[], 'SKLADY':[]}    
        self.display()


    def go_left(self):
        if self.DISPLAYINDEX>0:
            self.DISPLAYINDEX-=1

        self.display() 


    def go_right(self):
        if self.DISPLAYINDEX<len(self.DISPLAYIMAGESLIST)-1:
            self.DISPLAYINDEX+=1

        self.display()


    def delete_one(self):
        if len(self.DISPLAYIMAGESLIST) != 0 and not self.IMAGESLIST == [PIL.Image.new('RGB', (self.A4w, self.A4h), color='white')]:
            white = PIL.Image.new('RGB', (self.A4w//2, self.hgrid), color='white')
            n = self.LAST_GRID[0]
            m = self.LAST_GRID[1]
            self.CURRENT_ON_GRID = self.LAST_GRID
            self.LAST_GRID = self.ORDER[self.ORDER.index(self.CURRENT_ON_GRID)-1]
            self.CURRENT_PAGE.paste(white, (472//10+n*self.A4w//2, m*self.hgrid))

            data = BytesIO()
            self.CURRENT_PAGE.save(data, format='png')
            data.seek(0)
            cur_display = CoreImage(BytesIO(data.read()), ext='png').texture


            if (n,m) == (0,0):
                self.DISPLAYIMAGESLIST.pop(-1)
                self.IMAGESLIST.pop(-1)
                self.DISPLAYINDEX-=1
                try:
                    self.CURRENT_PAGE = self.IMAGESLIST[self.DISPLAYINDEX]
                
                except IndexError:
                    self.CURRENT_PAGE = PIL.Image.new('RGB', (self.A4w, self.A4h), color='white')
                    self.LAST_GRID = (0,0)
            
            else:
                self.DISPLAYIMAGESLIST[-1] = cur_display

            # self.DISPLAYINDEX = len(self.DISPLAYIMAGESLIST)-1
            #self.DISPLAYINDEX = 0 if self.DISPLAYINDEX < 0 else self.DISPLAYINDEX
            self.display()


class QuickLabel(App):

    def build(self):
        return MyGrid()



if __name__ == '__main__':
    QuickLabel().run()
    try:
        os.remove(f'{os.getcwd()}/tmpQL.pdf')
    
    except:
        pass
