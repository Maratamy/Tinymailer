import gi
import imaplib
import time
import email
from email.header import decode_header
import base64
import pickle
import os
import datetime
import locale
import ssl
import logging


gi.require_version("Gtk", "3.0")
from gi.repository import Gtk, GLib, Gdk


def label_cutter(label):
    if len(label) > 50:
        return label[0:50:1] + "..."
    else:
        return label


def format_date(date_str):
    locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
    date_str = date_str.strip('"').split(' +')[0]
    date_obj = datetime.datetime.strptime(date_str, '%d-%b-%Y %H:%M:%S')
    formatted_date = date_obj.strftime('%d %B %Y %H:%M')
    return formatted_date


class DataProcessor:
    @staticmethod
    def get_dir():
        if os.name == 'nt':  # Windows
            dir_path = os.path.join(os.environ['USERPROFILE'], 'Documents')
        elif os.name == 'posix':  # Linux or MacOS
            dir_path = os.path.join(os.environ['HOME'], 'Documents')
        return dir_path

    @staticmethod
    def data_saver(check, name, imap):
        dir_path = DataProcessor.get_dir()
        file_name = 'mymailprof.dat'
        file_path = os.path.join(dir_path, file_name)
        info = [check, name, imap]
        try:
            with open(file_path, "wb") as f:
                pickle.dump(info, f)
                f.close()
        except:
            print("Ошибка сохранения данных")

    @staticmethod
    def settings_saver(freq, show, on_screen):
        file_name = 'mymailset.dat'
        dir_path = DataProcessor.get_dir()
        file_path = os.path.join(dir_path, file_name)
        info = [freq, show, on_screen]
        try:
            with open(file_path, "wb") as f:
                pickle.dump(info, f)
                f.close()
        except:
            print("Ошибка сохранения данных")

    @staticmethod
    def data_reader(name):
        dir_path = DataProcessor.get_dir()
        file_path = os.path.join(dir_path, name)
        try:
            with open(file_path, "rb") as f:
                p = list(pickle.load(f))
                f.close()
                return p
        except:
            print("Ошибка чтения данных")


class Error(Gtk.Window):
    def __init__(self, error, title="Ошибка"):
        Gtk.Window.__init__(self, title=title)
        self.set_default_size(120, 50)
        self.set_resizable(False)
        self.set_border_width(10)
        self.label = Gtk.Label(label=str(error))
        self.label.set_justify(Gtk.Justification.FILL)
        self.add(self.label)
        self.show_all()


class ErrorDialog(Gtk.Dialog):
    def __init__(self, parent, err_text):
        Gtk.Dialog.__init__(self, title="Ошибка", transient_for=parent, flags=0)
        self.set_default_size(120, 50)
        self.set_resizable(False)
        self.set_border_width(10)
        self.label = Gtk.Label(label=err_text)
        self.label.set_justify(Gtk.Justification.FILL)
        self.get_content_area().add(self.label)
        self.show_all()
        self.run()


class Settings(Gtk.Window):
    def __init__(self, widget, cur_freq, show_me, on_top):
        Gtk.Window.__init__(self, title="Настройки")

        self.set_border_width(10)
        self.set_resizable(False)
        self.hbox = Gtk.Box(spacing=10)
        self.lbox = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        self.rbox = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        self.rbox.set_homogeneous(True)
        self.lbox.set_homogeneous(True)
        self.hbox.pack_start(self.lbox, True, True, 0)
        self.hbox.pack_start(self.rbox, True, True, 0)

        label = Gtk.Label(label="Частота обновления, с")
        self.lbox.pack_start(label, True, True, 0)
        label = Gtk.Label(label="Показать непрочитанные")
        self.lbox.pack_start(label, True, True, 0)
        self.adj = Gtk.Adjustment(value=cur_freq, lower=1, upper=400,
                                  step_increment=1, page_increment=10, page_size=0)
        self.frequency_slider = Gtk.HScale()
        self.frequency_slider.set_adjustment(self.adj)
        self.frequency_slider.set_digits(0)
        self.rbox.pack_start(self.frequency_slider, False, False, 0)
        self.amount_checkbox = Gtk.CheckButton()
        if show_me:
            self.amount_checkbox.set_active(True)
        self.rbox.pack_start(self.amount_checkbox, True, True, 0)
        label = Gtk.Label(label="Оповещение всегда сверху")
        self.lbox.pack_start(label, True, True, 0)
        self.always_on_top = Gtk.CheckButton()
        if on_top:
            self.always_on_top.set_active(True)
        self.rbox.pack_start(self.always_on_top, True, True, 0)
        self.button = Gtk.Button(label="Применить")
        self.rbox.pack_start(self.button, True, True, 0)
        label = Gtk.Label(label=" ")
        self.lbox.pack_start(label, True, True, 0)
        self.add(self.hbox)
        self.show_all()
        self.button.connect("clicked", self.on_button_clicked)

    def on_button_clicked(self, widget):
        try:
            DataProcessor.settings_saver(self.frequency_slider.get_value(),
                                      self.amount_checkbox.get_active(), self.always_on_top.get_active())
        except Exception as e:
            ErrorDialog(self, "Ошибка: " + str(e))


class Notify(Gtk.Window):
    def __init__(self, amount):
        Gtk.Window.__init__(self, title="Оповещение")
        self.set_default_size(300, 100)
        self.set_border_width(10)
        self.set_resizable(False)
        self.hbox = Gtk.Box()
        self.label = Gtk.Label(label="Непрочитанных сообщений: " + amount)
        self.hbox.pack_start(self.label, True, True, 0)
        self.add(self.hbox)
        self.show_all()


class MailProcessor():

    def __init__(self):
        self.on_screen = ''
        self.username = ''
        self.user_password = ''
        self.imap_server = ''

    def login_mail(self, username, passwrd, server, freq, on_screen):
        self.on_screen = on_screen
        self.username = username
        self.user_password = passwrd
        self.imap_server = server
        encoding = "utf-8"
        self.imap = imaplib.IMAP4_SSL(self.imap_server)
        self.imap.login(self.username, self.user_password)
        self.imap.select("INBOX")
        self.unseen_msg_old = self.imap.uid('search', "unseen", "all")[1]
        self.unseen_msg_old = set(self.unseen_msg_old[0].decode(encoding).split(" "))
        GLib.timeout_add(freq, self.read_mail)

    def get_amount_unseen_msg(self):
        return str(len(self.unseen_msg_old))

    def reconnect_imap(self):
        try:
            self.imap = imaplib.IMAP4_SSL(self.imap_server)
            self.imap.login(self.username, self.user_password)
            self.imap.select("INBOX")
            self.unseen_msg_old = self.imap.uid('search', "unseen", "all")[1]
            self.unseen_msg_old = set(self.unseen_msg_old[0].decode("utf-8").split(" "))
            return True
        except Exception as e:
            logging.error("Error occurred during reconnection: %s", e)
            time.sleep(3)
            return self.reconnect_imap()


    def read_mail(self):
        try:
            encoding = "utf-8"
            unseen_msg = self.imap.uid('search', "unseen", "all")[1]
            logging.info("Success")
            unseen_msg = set(unseen_msg[0].decode(encoding).split(" "))
            if unseen_msg - set(self.unseen_msg_old):
                letters = list(unseen_msg - set(self.unseen_msg_old))
                self.unseen_msg_old = unseen_msg
                for letter in letters:
                    res, msg = self.imap.uid('fetch', letter, '(RFC822)')
                    if res == "OK":
                        msg = email.message_from_bytes(msg[0][1])
                        msg_date = imaplib.Time2Internaldate(email.utils.parsedate_tz(msg["Date"]))
                        msg_from = msg["Return-path"]
                        msg_subj = decode_header(msg["Subject"])[0][0]
                        if isinstance(msg_subj, bytes):
                            msg_subj = msg_subj.decode()
                        formatted_date = format_date(msg_date)
                        NotificationWindow(formatted_date, msg_from, msg_subj, self.on_screen)
                        self.imap.uid('STORE', letter, '-FLAGS', '\\Seen')
                        self.imap.expunge()
            return True
        except Exception as e:
            logging.error("Error occurred: %s", e)
            print(self, "Ошибка: " + str(e))
            self.reconnect_imap()
            return True

class NotificationWindow(Gtk.Window):
    def __init__(self, Date, Path, Subj, Keep_Above):
        Path = str(Path).strip("<>").replace("<", "")
        Gtk.Window.__init__(self, title="Новое сообщение.")

        self.date = Gtk.Label(label=str(Date), expand=True, justify = Gtk.Justification.CENTER)
        self.path = Gtk.Label(label=label_cutter(str(Path)), expand=True, justify = Gtk.Justification.CENTER)
        self.subj = Gtk.Label(label=label_cutter(str(Subj)), expand=True)
        self.set_default_size(400, 100)
        self.set_border_width(10)
        self.grid = Gtk.Grid()

        display = Gdk.Display.get_default()
        monitor = display.get_primary_monitor()
        geometry = monitor.get_geometry()
        self.screen_width = geometry.width
        self.screen_height = geometry.height

        self.dialog_width = self.get_size()[0]
        self.dialog_height = self.get_size()[1]
        self.x = self.screen_width - self.dialog_width
        self.y = self.screen_height - self.dialog_height
        if Keep_Above:
            self.set_keep_above(True)
        label = Gtk.Label(label="Дата:")
        label.set_justify(Gtk.Justification.FILL)
        label.set_padding(0,0)
        self.grid.add(self.date)
        self.grid.attach_next_to(label, self.date, Gtk.PositionType.LEFT, 1, 1)
        label1 = Gtk.Label(label="Отправитель:")
        self.grid.attach_next_to(label1, label, Gtk.PositionType.BOTTOM, 1, 1)
        self.grid.attach_next_to(self.path, self.date, Gtk.PositionType.BOTTOM, 1, 1)
        self.grid.attach_next_to(self.subj, label1, Gtk.PositionType.BOTTOM, 2, 3)
        self.add(self.grid)
        self.set_resizable(False)
        self.move(self.x, self.y)
        self.show_all()


class MainWindow(Gtk.Window):

    def __init__(self):
        super().__init__(title="Tinymail")
        self.set_size_request(400, 300)
        self.set_resizable(False)
        self.set_border_width(10)

        self.status_icon = Gtk.StatusIcon()
        self.status_icon.set_from_icon_name("mail-inbox")
        self.status_icon.set_visible(True)
        self.connect("window-state-event", self.on_window_state_event)

        self.user_name_entry = Gtk.Entry()
        self.user_pass_entry = Gtk.Entry()
        self.user_pass_entry.set_visibility(False)
        self.user_pass_entry.set_invisible_char('•')
        self.user_imap_entry = Gtk.Entry()

        self.button_next = Gtk.Button(label="Далее")
        self.button_next.connect("clicked", self.on_button_clicked)
        self.status_icon.connect("button-press-event", self.on_status_icon_button_press)
        self.menu = Gtk.Menu()

        menu_item_expand = Gtk.MenuItem(label="Развернуть")
        menu_item_expand.connect("activate", self.on_status_icon_activate)
        self.menu.append(menu_item_expand)
        menu_item_expand.set_visible(True)

        menu_item_settings = Gtk.MenuItem(label="Настройки")
        menu_item_settings.connect("activate", self.on_settings_activated)
        self.menu.append(menu_item_settings)
        menu_item_settings.set_visible(True)

        menu_item_quit = Gtk.MenuItem(label="Выйти")
        menu_item_quit.connect("activate", Gtk.main_quit)
        self.menu.append(menu_item_quit)
        menu_item_quit.set_visible(True)

        self.check = False
        hbox = Gtk.Box(spacing=10)
        hbox.set_homogeneous(True)
        vbox_left = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        vbox_left.set_homogeneous(True)
        vbox_right = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        vbox_right.set_homogeneous(True)

        hbox.pack_start(vbox_left, True, True, 0)
        hbox.pack_start(vbox_right, True, True, 0)

        label = Gtk.Label(label="Имя пользователя")
        vbox_left.pack_start(label, True, True, 0)
        label = Gtk.Label()
        label.set_text("Пароль")
        label.set_justify(Gtk.Justification.LEFT)
        vbox_left.pack_start(label, True, True, 0)

        label = Gtk.Label(label="Имя imap сервера")
        label.set_justify(Gtk.Justification.LEFT)
        vbox_left.pack_start(label, True, True, 0)
        label = Gtk.Label(label=" ")
        vbox_left.pack_start(label, True, True, 0)
        label = Gtk.Label(label=" ")
        label.set_line_wrap(True)
        label.set_max_width_chars(32)
        vbox_right.pack_start(self.user_name_entry, True, True, 0)
        vbox_left.pack_start(label, True, True, 0)
        vbox_right.pack_start(self.user_pass_entry, True, True, 0)
        vbox_right.pack_start(self.user_imap_entry, True, True, 0)
        checkbox = Gtk.CheckButton(label="Запомнить меня")
        checkbox.connect("toggled", self.on_checkbox_toggled)
        vbox_right.pack_start(checkbox, True, True, 0)
        dir_path = DataProcessor.get_dir()
        if os.path.isfile(os.path.join(dir_path, "mymailprof.dat")):
            self.data = DataProcessor.data_reader("mymailprof.dat")
            if self.data[0]:
                self.check = True
                self.user_name_entry.set_text(self.data[1])
                self.user_imap_entry.set_text(self.data[2])
                checkbox.set_active(True)
        if os.path.isfile(os.path.join(dir_path, "mymailset.dat")):
            self.data = DataProcessor.data_reader("mymailset.dat")
            if self.data[0]:
                self.freq = self.data[0] * 1000
                self.show_me = self.data[1]
                self.on_screen = self.data[2]
        else:
            self.freq = 6000
            self.show_me = False
            self.on_screen = True
        vbox_right.pack_start(self.button_next, True, True, 0)
        self.add(hbox)

    def on_window_state_event(self, window, event):
        if event.new_window_state & Gdk.WindowState.ICONIFIED:
            self.hide()
        return False

    def on_settings_activated(self, widget):
        settings_window = Settings(self, self.freq / 1000, self.show_me, self.on_screen)
        settings_window.show()

    def on_settings_pressed(self, widget):
        Settings(widget=widget, cur_freq=self.freq/1000, show=self.show_me, on_top=self.on_screen)

    def on_checkbox_toggled(self, widget):
        self.check = widget.get_active()

    def on_status_icon_button_press(self, icon, event):
        if event.button == 3:
            self.menu.popup(None, None, func=None, button=event.button, activate_time=event.time, data=None)

    def on_status_icon_activate(self, status_icon):
        self.show()
        self.present()

    def on_window_destroy(self, widget, event):
        self.hide()
        self.status_icon.set_visible(True)
        self.disconnect("delete-event")
        Gtk.main_quit()

    def on_button_clicked(self, widget):
        try:
            mailer = MailProcessor()
            if self.check:
                DataProcessor.data_saver(self.check,
                                          str(self.user_name_entry.get_text()),
                                          str(self.user_imap_entry.get_text()))
            elif not self.check and os.path.isfile("profile.dat"):
                os.remove("profile.dat")
            mailer.login_mail(str(self.user_name_entry.get_text()), str(self.user_pass_entry.get_text()),
                              str(self.user_imap_entry.get_text()), self.freq, self.on_screen)
            if self.show_me:
                Notify(mailer.get_amount_unseen_msg)
            self.hide()
        except imaplib.IMAP4.error as e:
            ErrorDialog(self, "Ошибка авторизации. Проверьте введённые данные")
        except Exception as e:
            ErrorDialog(self, "Ошибка: " + str(e))


def main():
    win = MainWindow()
    win.connect("destroy", Gtk.main_quit)
    win.show_all()
    Gtk.main()


if __name__ == "__main__":
    try:
        main()
    except Exception as exp:
        text = str("Ошибка main: " + str(exp))
        print(text)
