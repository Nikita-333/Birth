import threading
import schedule
import time
import openpyxl
import plyer
import schedule
from kivy.config import Config
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.textinput import TextInput
from kivy.app import App
from kivy.core.window import Window

# Window.fullscreen = 'auto'
Config.set("graphics", "resizable", 0)
Config.set("graphics", "width", 500)
Config.set("graphics", "height", 500)
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
import kivymd
from kivymd.app import MDApp
from check import find_nearest_birthday, notify_birthday, run_scheduler, find_birthday_by_lastname, \
    add_employee_to_excel, delete_employee_by_name, find_next_birthday


class FileChooserPopup(Popup):
    def __init__(self, on_file_selected, **kwargs):
        super(FileChooserPopup, self).__init__(**kwargs)
        self.title = "Выберите файл"

        self.file_chooser = FileChooserListView(filters=['*.xlsx'])
        self.file_chooser.bind(on_submit=self.dismiss)

        self.add_widget(self.file_chooser)

        self.on_file_selected = on_file_selected

    def dismiss(self, *args):
        if self.file_chooser.selection:
            selected_file = self.file_chooser.selection[0]
            self.on_file_selected(selected_file)

        super(FileChooserPopup, self).dismiss()


class BirthdayApp(App):
    def build(self):
        self.screen_manager = ScreenManager()

        main_screen = Screen(name='main')
        layout = BoxLayout(orientation='vertical', spacing=10)
        main_screen.add_widget(layout)

        self.selected_file_label = Label(size_hint=(1, None), height=50)
        layout.add_widget(self.selected_file_label)

        self.result_label = Label(size_hint=(1, None))
        self.result_label.bind(texture_size=self.result_label.setter('size'))
        layout.add_widget(self.result_label)

        buttons_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), size_hint_y=None)

        select_file_button = Button(text="Выбрать файл", size_hint=(1, None), height=50)
        select_file_button.bind(on_release=self.show_file_chooser)
        buttons_layout.add_widget(select_file_button)

        nearest_birthday_button = Button(text='Ближайшее день рождение', size_hint=(1, None), height=50)
        nearest_birthday_button.bind(on_release=self.on_nearest_birthday_button_release)
        buttons_layout.add_widget(nearest_birthday_button)

        layout.add_widget(buttons_layout)

        search_by_lastname_button = Button(text='Поиск по фамилии', size_hint=(1, None), height=50)
        search_by_lastname_button.bind(on_release=self.switch_to_search_screen)
        layout.add_widget(search_by_lastname_button)

        add_employee_button = Button(text='Добавить сотрудника', size_hint=(1, None), height=50)
        add_employee_button.bind(on_release=self.switch_to_add_screen)
        layout.add_widget(add_employee_button)

        delete_employee_button = Button(text='Удалить сотрудника', size_hint=(1, None), height=50)
        delete_employee_button.bind(on_release=self.switch_to_del_screen)
        layout.add_widget(delete_employee_button)

        self.screen_manager.add_widget(main_screen)

        search_screen = Screen(name='search')
        search_layout = BoxLayout(orientation='vertical', spacing=10)
        search_screen.add_widget(search_layout)

        self.search_results_label = Label(size_hint=(1, None))
        self.search_results_label.bind(texture_size=self.search_results_label.setter('size'))
        search_layout.add_widget(self.search_results_label)

        search_input = TextInput(hint_text='Введите фамилию', size_hint=(1, None), height=50)
        search_layout.add_widget(search_input)

        search_button = Button(text='Подтвердить', size_hint=(1, None), height=50)
        search_button.bind(on_release=lambda instance: self.show_search_results(search_input.text))
        search_layout.add_widget(search_button)

        back_button = Button(text='Назад', size_hint=(1, None), height=50)
        back_button.bind(on_release=self.switch_to_main_screen)
        search_layout.add_widget(back_button)

        self.screen_manager.add_widget(search_screen)

        add_employee_screen = Screen(name='add')
        add_layout = BoxLayout(orientation='vertical', spacing=10)
        add_employee_screen.add_widget(add_layout)

        self.add_results_label = Label(size_hint=(1, None))
        self.add_results_label.bind(texture_size=self.add_results_label.setter('size'))
        add_layout.add_widget(self.add_results_label)

        add_input = TextInput(hint_text='Введите ФИО', size_hint=(1, None), height=50)
        add_layout.add_widget(add_input)

        add_input1 = TextInput(hint_text='Введите дату рождения в формате MM/DD/YY', size_hint=(1, None), height=50)
        add_layout.add_widget(add_input1)

        add_button = Button(text='Подтвердить', size_hint=(1, None), height=50)
        add_button.bind(on_release=lambda instance: self.add_employee(add_input.text, add_input1.text))
        add_layout.add_widget(add_button)

        back_button = Button(text='Назад', size_hint=(1, None), height=50)
        back_button.bind(on_release=self.switch_to_main_screen)
        add_layout.add_widget(back_button)

        self.screen_manager.add_widget(add_employee_screen)

        del_employee_screen = Screen(name='del')
        del_layout = BoxLayout(orientation='vertical', spacing=10)
        del_employee_screen.add_widget(del_layout)

        self.del_results_label = Label(size_hint=(1, None))
        self.del_results_label.bind(texture_size=self.del_results_label.setter('size'))
        del_layout.add_widget(self.del_results_label)

        del_input = TextInput(hint_text='Введите ФИО сотрудника', size_hint=(1, None), height=50)
        del_layout.add_widget(del_input)

        del_button = Button(text='Подтвердить', size_hint=(1, None), height=50)
        del_button.bind(on_release=lambda instance: self.del_employee(del_input.text))
        del_layout.add_widget(del_button)

        back_button = Button(text='Назад', size_hint=(1, None), height=50)
        back_button.bind(on_release=self.switch_to_main_screen)
        del_layout.add_widget(back_button)

        self.screen_manager.add_widget(del_employee_screen)

        return self.screen_manager

    def show_file_chooser(self, *args):
        file_chooser_popup = FileChooserPopup(on_file_selected=self.on_file_selected)
        file_chooser_popup.open()

    def on_file_selected(self, selected_file):
        self.selected_file_label.text = f"Выбранный файл: {selected_file}"

    def on_nearest_birthday_button_release(self, instance):
        selected_file = self.selected_file_label.text.replace("Выбранный файл: ", "")
        result = find_nearest_birthday(selected_file)
        self.result_label.text = result if result is not None else "Нет ближайших дней рождений."

    def on_window_resize(self, window, width, height):
        self.result_label.text_size = (self.result_label.width, None)

    def switch_to_search_screen(self, *args):
        self.screen_manager.current = 'search'

    def switch_to_main_screen(self, *args):
        self.screen_manager.current = 'main'

    def switch_to_add_screen(self, *args):
        self.screen_manager.current = 'add'

    def switch_to_del_screen(self, *args):
        self.screen_manager.current = 'del'

    def show_search_results(self, name_or_lastname):
        selected_file = self.selected_file_label.text.replace("Выбранный файл: ", "")
        result = find_birthday_by_lastname(selected_file, name_or_lastname)
        self.search_results_label.text = result if result else "Нет сотрудников"
        self.screen_manager.current = 'search'

    def find_birthday(self, name_or_lastname):
        selected_file = self.selected_file_label.text.replace("Выбранный файл: ", "")
        result = find_birthday_by_lastname(selected_file, name_or_lastname)
        self.result_label.text = result if result else "Нет сотрудников"

    def add_employee(self, full_name, date_of_birth):
        selected_file = self.selected_file_label.text.replace("Выбранный файл: ", "")
        result = add_employee_to_excel(selected_file, full_name, date_of_birth)
        self.add_results_label.text = result if result else "Невозможно добавить"
        self.screen_manager.current = 'add'

    def del_employee(self, full_name):
        selected_file = self.selected_file_label.text.replace("Выбранный файл: ", "")
        result = delete_employee_by_name(selected_file, full_name)
        self.del_results_label.text = result if result else "Невозможно удалить"
        self.screen_manager.current = 'del'
