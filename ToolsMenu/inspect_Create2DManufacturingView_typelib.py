# -*- coding: cp1251 -*-
"""
Инспектор COM-типобиблиотек E3.series для поиска Create2DManufacturingView

Запуск: используйте python из проекта venv или добавьте в Tools -> Add-ons.
Скрипт попытается подключиться к CT.Application, получить содержащую TypeLib
и просканировать все TypeInfo на наличие имени Create2DManufacturingView.

Если метод найден — будут выведены имена параметров и доступная документация.
"""

import sys
import traceback

try:
    import win32com.client
    import pythoncom
except Exception as e:
    print("Требуются pywin32 (win32com, pythoncom). Установите пакет pywin32 в venv:")
    print("pip install pywin32")
    sys.exit(2)

TARGET_NAME = "Create2DManufacturingView"


def search_in_typelib(typelib):
    found = []
    try:
        count = typelib.GetTypeInfoCount()
    except Exception as e:
        print("Ошибка получения количества TypeInfo:", e)
        return found

    for i in range(count):
        try:
            tinfo = typelib.GetTypeInfo(i)
        except Exception:
            continue
        try:
            type_doc = tinfo.GetDocumentation(-1)
        except Exception:
            type_doc = ("", "", 0, "")
        try:
            tattr = tinfo.GetTypeAttr()
        except Exception:
            continue
        cfuncs = getattr(tattr, 'cFuncs', 0)
        for fi in range(cfuncs):
            try:
                fd = tinfo.GetFuncDesc(fi)
            except Exception:
                continue
            memid = getattr(fd, 'memid', None)
            if memid is None:
                continue
            try:
                names = tinfo.GetNames(memid)
            except Exception:
                names = ()
            # names: (funcname, param1, param2, ..., optional help)
            if any(n == TARGET_NAME for n in names):
                # get member documentation if available
                try:
                    member_doc = tinfo.GetDocumentation(memid)
                except Exception:
                    member_doc = ("", "", 0, "")
                found.append({
                    'type_index': i,
                    'type_name': type_doc[0],
                    'type_doc': type_doc[1],
                    'func_index': fi,
                    'names': names,
                    'member_doc': member_doc,
                })
    return found


def main():
    try:
        # Подключимся к приложению (если E3 не запущен, это может создать инстанс)
        try:
            app = win32com.client.Dispatch("CT.Application")
        except Exception as e:
            print("Не удалось подключиться к CT.Application:", e)
            print("Попробуйте запустить E3.series перед запуском скрипта или запустите makepy для зарегистрированной TLB.")
            return 3

        try:
            typeinfo = app._oleobj_.GetTypeInfo()
            typelib, idx = typeinfo.GetContainingTypeLib()
        except Exception as e:
            print("Не удалось получить TypeLib от CT.Application:", e)
            traceback.print_exc()
            return 4

        print("TypeLib index внутри контейнера:", idx)
        results = search_in_typelib(typelib)
        if not results:
            print(f"{TARGET_NAME} не найден в TypeLib, привязанной к запущенному E3.")
            print("Если у вас старая версия E3 — метод физически может отсутствовать в COM.")
            return 0

        for r in results:
            print("--- Найдено ---")
            print("TypeIndex:", r['type_index'])
            print("TypeName:", r['type_name'])
            print("FuncIndex:", r['func_index'])
            print("GetNames:", r['names'])
            print("Member doc:")
            print(r['member_doc'])

        return 0

    except Exception as exc:
        print("Фатальная ошибка:", exc)
        traceback.print_exc()
        return 2


if __name__ == '__main__':
    sys.exit(main())
