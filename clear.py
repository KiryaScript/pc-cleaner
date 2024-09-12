import os
import shutil
import send2trash
import hashlib
import winreg
import subprocess
import json
import threading
import concurrent.futures
import time
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageTk
from collections import defaultdict
import psutil

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AnimatedGIF(ctk.CTkLabel):
    def __init__(self, master, path, size=(100, 100)):
        self.frames = []
        self.current_frame = 0
        
        image = Image.open(path)
        for frame in range(image.n_frames):
            image.seek(frame)
            frame_image = image.copy().resize(size, Image.LANCZOS)
            self.frames.append(ImageTk.PhotoImage(frame_image))
        
        super().__init__(master, image=self.frames[0], text="")
        self.is_playing = False

    def start(self):
        self.is_playing = True
        self.animate()

    def stop(self):
        self.is_playing = False

    def animate(self):
        if self.is_playing:
            self.current_frame = (self.current_frame + 1) % len(self.frames)
            self.configure(image=self.frames[self.current_frame])
            self.after(50, self.animate)

class PCCleaner:
    def __init__(self, log_widget):
        self.log_widget = log_widget
        self.stop_flag = threading.Event()
        self.report_data = {}

    def log(self, message):
        self.log_widget.insert("end", message + "\n")
        self.log_widget.see("end")
        self.report_data[time.strftime("%Y-%m-%d %H:%M:%S")] = message

    def analyze_installed_programs(self):
        self.log("Анализ установленных программ...")
        programs = self.get_installed_programs()
        self.show_installed_programs(programs)

    def get_installed_programs(self):
        programs = []
        for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
            try:
                key = winreg.OpenKey(hive, key_path)
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                        programs.append((display_name, uninstall_string))
                    except WindowsError:
                        continue
            except WindowsError:
                continue
        return programs

    def show_installed_programs(self, programs):
        programs_window = ctk.CTkToplevel()
        programs_window.title("Установленные программы")
        programs_window.geometry("600x400")

        listbox = ctk.CTkTextbox(programs_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for program, _ in sorted(programs):
            listbox.insert("end", f"{program}\n")

        uninstall_button = ctk.CTkButton(programs_window, text="Удалить выбранную программу", 
                                        command=lambda: self.uninstall_program(listbox, programs))
        uninstall_button.pack(pady=10)

    def uninstall_program(self, listbox, programs):
        selected = listbox.selection_get().split("\n")[0]
        for program_name, uninstall_string in programs:
            if program_name == selected:
                if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить {program_name}?"):
                    try:
                        subprocess.run(uninstall_string, shell=True, check=True)
                        self.log(f"Программа {program_name} успешно удалена")
                    except subprocess.CalledProcessError:
                        self.log(f"Ошибка при удалении программы {program_name}")
                break

    def export_report(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.report_data, f, ensure_ascii=False, indent=4)
            self.log(f"Отчет экспортирован в {file_path}")

    def analyze_large_files(self):
        path = filedialog.askdirectory(title="Выберите директорию для анализа больших файлов")
        if not path:
            return

        self.log(f"Анализ больших файлов в {path}")
        large_files = []
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            for root, _, files in os.walk(path):
                if self.stop_flag.is_set():
                    break
                futures = [executor.submit(self.get_file_size, os.path.join(root, file)) for file in files]
                for future in concurrent.futures.as_completed(futures):
                    if self.stop_flag.is_set():
                        break
                    file_path, size = future.result()
                    if size > 100 * 1024 * 1024:  # файлы больше 100 МБ
                        large_files.append((file_path, size))

        large_files.sort(key=lambda x: x[1], reverse=True)
        self.show_large_files(large_files[:100])  # показываем топ-100 больших файлов

    def show_large_files(self, large_files):
        large_files_window = ctk.CTkToplevel()
        large_files_window.title("Большие файлы")
        large_files_window.geometry("600x400")

        listbox = ctk.CTkTextbox(large_files_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for file_path, size in large_files:
            file_name = os.path.basename(file_path)
            listbox.insert("end", f"{self.format_size(size)} - {file_name}\n")

        delete_button = ctk.CTkButton(large_files_window, text="Удалить выбранный файл", 
                                    command=lambda: self.delete_large_file(listbox, large_files))
        delete_button.pack(pady=10)

    def manage_startup_programs(self):
        startup_programs = self.get_startup_programs()
        self.show_startup_programs(startup_programs)

    def get_startup_programs(self):
        startup_programs = []
        startup_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        
        # Проверяем программы в папке автозагрузки
        for item in os.listdir(startup_path):
            item_path = os.path.join(startup_path, item)
            startup_programs.append((item, item_path))
        
        # Проверяем программы в реестре
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        for hive in [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]:
            try:
                key = winreg.OpenKey(hive, key_path)
                for i in range(winreg.QueryInfoKey(key)[1]):
                    name, value, type = winreg.EnumValue(key, i)
                    startup_programs.append((name, value))
            except WindowsError:
                continue
        
        return startup_programs

    def show_startup_programs(self, startup_programs):
        startup_window = ctk.CTkToplevel()
        startup_window.title("Программы автозагрузки")
        startup_window.geometry("600x400")

        listbox = ctk.CTkTextbox(startup_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for program, path in startup_programs:
            listbox.insert("end", f"{program}\n")

        disable_button = ctk.CTkButton(startup_window, text="Отключить выбранную программу", 
                                   command=lambda: self.disable_startup_program(listbox, startup_programs))
        disable_button.pack(pady=10)

    def disable_startup_program(self, listbox, startup_programs):
        selected = listbox.selection_get().split("\n")[0]
        for program, path in startup_programs:
            if program == selected:
                if path.endswith('.lnk'):
                    try:
                        os.remove(path)
                        self.log(f"Программа {program} удалена из автозагрузки")
                    except Exception as e:
                        self.log(f"Ошибка при удалении {program} из автозагрузки: {str(e)}")
                else:
                    try:
                        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
                        winreg.DeleteValue(key, program)
                        winreg.CloseKey(key)
                        self.log(f"Программа {program} удалена из автозагрузки")
                    except WindowsError as e:
                        self.log(f"Ошибка при удалении {program} из автозагрузки: {str(e)}")
                    break
        
        # Обновляем список программ автозагрузки
        updated_startup_programs = self.get_startup_programs()
        listbox.delete("1.0", "end")
        for program, _ in updated_startup_programs:
            listbox.insert("end", f"{program}\n")

    def clean_temp_files(self):
        temp_folders = [os.environ.get('TEMP'), os.environ.get('TMP')]
        for folder in temp_folders:
            if folder and os.path.exists(folder):
                self.log(f"Очистка временной папки: {folder}")
                with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
                    for root, dirs, files in os.walk(folder):
                        if self.stop_flag.is_set():
                            break
                        futures = []
                        for name in files + dirs:
                            path = os.path.join(root, name)
                            if os.path.isfile(path):
                                futures.append(executor.submit(self.delete_file, path))
                            elif os.path.isdir(path):
                                futures.append(executor.submit(self.delete_directory, path))
                        concurrent.futures.wait(futures)
        self.log("Очистка временных файлов завершена")

    def delete_file(self, path):
        try:
            os.unlink(path)
            self.log(f"Удален файл: {path}")
        except Exception as e:
            self.log(f"Ошибка при удалении {path}: {str(e)}")

    def delete_directory(self, path):
        try:
            shutil.rmtree(path)
            self.log(f"Удалена папка: {path}")
        except Exception as e:
            self.log(f"Ошибка при удалении {path}: {str(e)}")

    def empty_recycle_bin(self):
        self.log("Начало очистки корзины")
        try:
            send2trash.send2trash = lambda x: x
            shutil.rmtree(os.path.expanduser("~/$Recycle.Bin"))
            self.log("Корзина успешно очищена")
        except Exception as e:
            self.log(f"Ошибка при очистке корзины: {str(e)}")
        finally:
            send2trash.send2trash = send2trash._send2trash

    def find_duplicates(self):
        path = filedialog.askdirectory(title="Выберите директорию для поиска дубликатов")
        if not path:
            return

        self.log(f"Начало поиска дубликатов в {path}")
        duplicates = defaultdict(list)
        
        start_time = time.time()
        total_size = 0
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            futures = []
            for dirpath, _, filenames in os.walk(path):
                for filename in filenames:
                    if self.stop_flag.is_set():
                        break
                    full_path = os.path.join(dirpath, filename)
                    futures.append(executor.submit(self.process_file, full_path))
            
            for future in concurrent.futures.as_completed(futures):
                if self.stop_flag.is_set():
                    break
                file_hash, file_path, file_size = future.result()
                duplicates[file_hash].append(file_path)
                total_size += file_size

        end_time = time.time()
        elapsed_time = end_time - start_time
        speed = total_size / elapsed_time / (1024 * 1024)  # MB/s

        self.log(f"Обработано {total_size / (1024*1024):.2f} MB за {elapsed_time:.2f} секунд")
        self.log(f"Средняя скорость: {speed:.2f} MB/s")

        duplicates = {k: v for k, v in duplicates.items() if len(v) > 1}
        if duplicates:
            self.show_duplicates(duplicates)
        else:
            self.log("Дубликаты не найдены")

    def process_file(self, path):
        file_hash = self.hash_file(path)
        file_size = os.path.getsize(path)
        return file_hash, path, file_size

    def hash_file(self, path):
        h = hashlib.sha1()
        with open(path, 'rb') as file:
            while chunk := file.read(1024 * 1024):  # Read 1MB at a time
                h.update(chunk)
        return h.hexdigest()

    def show_duplicates(self, duplicates):
        dup_window = ctk.CTkToplevel()
        dup_window.title("Найденные дубликаты")
        dup_window.geometry("500x400")

        listbox = ctk.CTkTextbox(dup_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for files in duplicates.values():
            listbox.insert("end", "Группа дубликатов:\n")
            for file in files:
                listbox.insert("end", f"{file}\n")
            listbox.insert("end", "\n")

        delete_button = ctk.CTkButton(dup_window, text="Удалить выбранные", 
                                      command=lambda: self.delete_selected_duplicates(listbox, duplicates))
        delete_button.pack(pady=10)

    def delete_selected_duplicates(self, listbox, duplicates):
        selected = listbox.selection_get().split("\n")
        to_delete = [file for file in selected if os.path.isfile(file)]

        if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить {len(to_delete)} файлов?"):
            with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
                futures = [executor.submit(self.delete_file, file) for file in to_delete]
                concurrent.futures.wait(futures)
            self.log(f"Удалено {len(to_delete)} файлов")

    def analyze_disk_space(self):
        path = filedialog.askdirectory(title="Выберите директорию для анализа")
        if not path:
            return

        self.log(f"Начало анализа дискового пространства в {path}")
        total_size = 0
        file_sizes = {}
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            futures = []
            for dirpath, _, filenames in os.walk(path):
                for filename in filenames:
                    if self.stop_flag.is_set():
                        break
                    full_path = os.path.join(dirpath, filename)
                    futures.append(executor.submit(self.get_file_size, full_path))
            
            for future in concurrent.futures.as_completed(futures):
                if self.stop_flag.is_set():
                    break
                path, size = future.result()
                total_size += size
                file_sizes[path] = size

        self.show_disk_analysis(total_size, file_sizes)

    def get_file_size(self, path):
        try:
            size = os.path.getsize(path)
            return path, size
        except Exception as e:
            self.log(f"Ошибка при получении размера файла {path}: {str(e)}")
            return path, 0

    def show_disk_analysis(self, total_size, file_sizes):
        analysis_window = ctk.CTkToplevel()
        analysis_window.title("Анализ дискового пространства")
        analysis_window.geometry("500x400")

        ctk.CTkLabel(analysis_window, text=f"Общий размер: {self.format_size(total_size)}").pack(pady=10)

        listbox = ctk.CTkTextbox(analysis_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for path, size in sorted(file_sizes.items(), key=lambda x: x[1], reverse=True)[:100]:
            listbox.insert("end", f"{self.format_size(size)} - {path}\n")

    def format_size(self, size):
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0

    def analyze_installed_programs(self):
        self.log("Анализ установленных программ...")
        programs = self.get_installed_programs()
        self.show_installed_programs(programs)

    def get_installed_programs(self):
        programs = []
        for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
            try:
                key = winreg.OpenKey(hive, key_path)
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                        programs.append((display_name, uninstall_string))
                    except WindowsError:
                        continue
            except WindowsError:
                continue
        return programs

    def show_installed_programs(self, programs):
        programs_window = ctk.CTkToplevel()
        programs_window.title("Установленные программы")
        programs_window.geometry("600x400")

        listbox = ctk.CTkTextbox(programs_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for program, _ in sorted(programs):
            listbox.insert("end", f"{program}\n")

        uninstall_button = ctk.CTkButton(programs_window, text="Удалить выбранную программу", 
                                         command=lambda: self.uninstall_program(listbox, programs))
        uninstall_button.pack(pady=10)

    def uninstall_program(self, listbox, programs):
        selected = listbox.selection_get().split("\n")[0]
        for program_name, uninstall_string in programs:
            if program_name == selected:
                if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить {program_name}?"):
                    try:
                        subprocess.run(uninstall_string, shell=True, check=True)
                        self.log(f"Программа {program_name} успешно удалена")
                    except subprocess.CalledProcessError:
                        self.log(f"Ошибка при удалении программы {program_name}")
                break

    def export_report(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.report_data, f, ensure_ascii=False, indent=4)
            self.log(f"Отчет экспортирован в {file_path}")

    def analyze_large_files(self):
        path = filedialog.askdirectory(title="Выберите директорию для анализа больших файлов")
        if not path:
            return

        self.log(f"Анализ больших файлов в {path}")
        large_files = []
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            for root, _, files in os.walk(path):
                if self.stop_flag.is_set():
                    break
                futures = [executor.submit(self.get_file_size, os.path.join(root, file)) for file in files]
                for future in concurrent.futures.as_completed(futures):
                    if self.stop_flag.is_set():
                        break
                    file_path, size = future.result()
                    if size > 100 * 1024 * 1024:  # файлы больше 100 МБ
                        large_files.append((file_path, size))

        large_files.sort(key=lambda x: x[1], reverse=True)
        self.show_large_files(large_files[:100])  # показываем топ-100 больших файлов

    def show_large_files(self, large_files):
        large_files_window = ctk.CTkToplevel()
        large_files_window.title("Большие файлы")
        large_files_window.geometry("600x400")

        listbox = ctk.CTkTextbox(large_files_window)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for file_path, size in large_files:
            listbox.insert("end", f"{self.format_size(size)} - {file_path}\n")

        delete_button = ctk.CTkButton(large_files_window, text="Удалить выбранный файл", 
                                      command=lambda: self.delete_large_file(listbox, large_files))
        delete_button.pack(pady=10)

    def delete_large_file(self, listbox, large_files):
        selected = listbox.selection_get().split("\n")[0]
        selected_size, selected_name = selected.split(" - ")
        for file_path, size in large_files:
            if os.path.basename(file_path) == selected_name and self.format_size(size) == selected_size:
                if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить файл {selected_name}?"):
                    try:
                        os.remove(file_path)
                        self.log(f"Файл {selected_name} успешно удален")
                        listbox.delete("1.0", "end")
                        for f_path, f_size in large_files:
                            if f_path != file_path:
                                f_name = os.path.basename(f_path)
                                listbox.insert("end", f"{self.format_size(f_size)} - {f_name}\n")
                    except Exception as e:
                        self.log(f"Ошибка при удалении файла {selected_name}: {str(e)}")
                    break

    def clean_and_optimize(self):
        self.clean_temp_files()
        self.empty_recycle_bin()
        self.log("Очистка и оптимизация завершены")

    def stop_operations(self):
        self.stop_flag.set()
        self.log("Операции остановлены пользователем")

class PCCleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Cum Cleaner")
        self.geometry("900x600")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Cum Cleaner v0.1", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.sidebar_button_1 = ctk.CTkButton(self.sidebar_frame, text="Очистка и оптимизация", command=self.clean_and_optimize)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)

        self.sidebar_button_2 = ctk.CTkButton(self.sidebar_frame, text="Анализ дубликатов", command=self.find_duplicates)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)

        self.sidebar_button_3 = ctk.CTkButton(self.sidebar_frame, text="Анализ диска", command=self.analyze_disk_space)
        self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=10)

        self.sidebar_button_4 = ctk.CTkButton(self.sidebar_frame, text="Анализ программ", command=self.analyze_installed_programs)
        self.sidebar_button_4.grid(row=4, column=0, padx=20, pady=10)

        self.sidebar_button_5 = ctk.CTkButton(self.sidebar_frame, text="Большие файлы", command=self.analyze_large_files)
        self.sidebar_button_5.grid(row=5, column=0, padx=20, pady=10)

        self.sidebar_button_6 = ctk.CTkButton(self.sidebar_frame, text="Управление автозагрузкой", command=self.manage_startup_programs)
        self.sidebar_button_6.grid(row=6, column=0, padx=20, pady=10)

        self.sidebar_button_7 = ctk.CTkButton(self.sidebar_frame, text="Экспорт отчета", command=self.export_report)
        self.sidebar_button_7.grid(row=7, column=0, padx=20, pady=10)

        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=9, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=10, column=0, padx=20, pady=(10, 10))

        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self.log_label = ctk.CTkLabel(self.main_frame, text="Лог операций:", anchor="w")
        self.log_label.grid(row=0, column=0, padx=20, pady=(20, 0), sticky="w")

        self.log_textbox = ctk.CTkTextbox(self.main_frame, width=200)
        self.log_textbox.grid(row=1, column=0, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.cleaner = PCCleaner(self.log_textbox)

        self.animated_gif = AnimatedGIF(self.sidebar_frame, "cat-girl.gif", size=(100, 100))
        self.animated_gif.grid(row=11, column=0, padx=20, pady=20)
        self.animated_gif.start()

    def clean_and_optimize(self):
        threading.Thread(target=self.cleaner.clean_and_optimize).start()

    def find_duplicates(self):
        threading.Thread(target=self.cleaner.find_duplicates).start()

    def analyze_disk_space(self):
        threading.Thread(target=self.cleaner.analyze_disk_space).start()

    def analyze_installed_programs(self):
        threading.Thread(target=self.cleaner.analyze_installed_programs).start()

    def analyze_large_files(self):
        threading.Thread(target=self.cleaner.analyze_large_files).start()

    def manage_startup_programs(self):
        threading.Thread(target=self.cleaner.manage_startup_programs).start()

    def export_report(self):
        threading.Thread(target=self.cleaner.export_report).start()

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

if __name__ == "__main__":
    app = PCCleanerApp()
    app.mainloop()