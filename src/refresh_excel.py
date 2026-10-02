import win32com.client as win32
import os
import time
import subprocess
import re
import sys
import argparse
from pathlib import Path
# === НАСТРОЙКИ ===
TIMEOUT_SECONDS = 600
EXCEL_VISIBLE = False # Для автоматизации лучше False, чтобы не мелькали окна
# === СКРИПТ ДЛЯ ОБРАБОТКИ ОДНОГО ФАЙЛА (SUBPROCESS) ===
SINGLE_FILE_SCRIPT = '''
import win32com.client as win32
import sys
import time
filepath = sys.argv[1]
visible = sys.argv[2] == "True"
try:
   excel = win32.gencache.EnsureDispatch('Excel.Application')
   excel.Visible = visible
   excel.DisplayAlerts = False
   excel.AskToUpdateLinks = False
   excel.EnableEvents = False
   wb = excel.Workbooks.Open(filepath, UpdateLinks=0, IgnoreReadOnlyRecommended=True)
   wb.RefreshAll()
   excel.CalculateUntilAsyncQueriesDone()
   time.sleep(2)
   wb.Save()
   wb.Close(SaveChanges=False)
   excel.Quit()
   print("SUCCESS")
except Exception as e:
   print(f"ERROR:{e}")
   try:
       excel.Quit()
   except:
       pass
'''
def normalize_apostrophes(text):
   return re.sub(r"[`´'']", "'", text)
def process_file_in_subprocess(filepath, timeout):
   try:
       process = subprocess.Popen(
           [sys.executable, "-c", SINGLE_FILE_SCRIPT, str(filepath), str(EXCEL_VISIBLE)],
           stdout=subprocess.PIPE,
           stderr=subprocess.PIPE,
           text=True
       )
       stdout, stderr = process.communicate(timeout=timeout)
       if "SUCCESS" in stdout:
           return True, None
       elif "ERROR:" in stdout:
           error_msg = stdout.split("ERROR:")[1].strip()
           return False, error_msg
       else:
           return False, f"Неизвестный результат: {stdout}"
   except subprocess.TimeoutExpired:
       process.kill()
       process.wait()
       return False, f"Таймаут {timeout}с"
   except Exception as e:
       return False, str(e)
def main():
   parser = argparse.ArgumentParser(description="Автоматическое обновление Excel файлов Performance Tracking (OLAP/ODC).")
   parser.add_argument("--path", type=str, help="Путь к папке с Excel-файлами")
   parser.add_argument("--month", type=str, default="", help="Фильтр по месяцу (например, May'26)")
   args = parser.parse_args()
   # Ввод данных (интерактивный или из аргументов)
   if args.path:
       pathname = args.path
   else:
       while True:
           pathname = input("Введите путь к папке с Excel-файлами: ").strip().strip('"').strip("'")
           if os.path.isdir(pathname):
               break
           print(f"Папка не найдена: {pathname}. Попробуйте ещё раз.\n")
   if args.month:
       month_filter = args.month
   else:
       month_filter = input("Введите месяц для поиска (например, May'26, или Enter для всех): ").strip()
   # Поиск файлов
   excel_extensions = ('.xlsx', '.xlsm', '.xls', '.xlsb')
   root_path = Path(pathname)
   all_files = [f for f in root_path.rglob('*')
                if f.is_file() and f.suffix.lower() in excel_extensions]
   if month_filter:
       normalized_filter = normalize_apostrophes(month_filter)
       files_to_process = [f for f in all_files
                          if normalized_filter in normalize_apostrophes(f.name)]
   else:
       files_to_process = all_files
   print(f"\nНайдено файлов всего: {len(all_files)}")
   print(f"Подходят по фильтру '{month_filter or 'все'}': {len(files_to_process)}")
   if not files_to_process:
       print("Нет файлов для обработки.")
       sys.exit(0)
   # Обработка файлов
   success_files = []
   failed_files = []
   for i, f in enumerate(files_to_process, 1):
       print(f"\n[{i}/{len(files_to_process)}] {f.name}")
       success, error = process_file_in_subprocess(f, TIMEOUT_SECONDS)
       if success:
           print("   Успешно")
           success_files.append(f.name)
       else:
           print(f"   ОШИБКА: {error}")
           failed_files.append((f.name, str(f.parent), error))
   # Итоговый отчет
   print("\n" + "="*70)
   print("ОБРАБОТКА ЗАВЕРШЕНА")
   print("="*70)
   print(f"Успешно обновлено: {len(success_files)}")
   print(f"С ошибками: {len(failed_files)}")
   print(f"Всего обработано: {len(success_files) + len(failed_files)}")
   if failed_files:
       print("\n" + "="*70)
       print("ФАЙЛЫ, ТРЕБУЮЩИЕ РУЧНОЙ ПРОВЕРКИ:")
       print("="*70)
       for idx, (name, path, reason) in enumerate(failed_files, 1):
           print(f"\n{idx}. {name}")
           print(f"   Путь: {path}")
           print(f"   Причина: {reason}")
       log_path = root_path / "failed_files_log.txt"
       try:
           with open(log_path, "w", encoding="utf-8") as log:
               log.write(f"Отчет от {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
               log.write(f"Успешно: {len(success_files)}, Ошибок: {len(failed_files)}\n\n")
               for name, path, reason in failed_files:
                   log.write(f"[!] {name}\n    Путь: {path}\n    Причина: {reason}\n\n")
           print(f"\nСписок проблемных файлов сохранен в: {log_path}")
       except Exception as e:
           print(f"\nНе удалось сохранить лог: {e}")
   print("="*70)
if __name__ == "__main__":
   main()
