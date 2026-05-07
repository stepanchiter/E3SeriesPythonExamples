# -*- coding: cp1251 -*-
"""
Создаёт 2D Manufacturing View для выделенного в дереве проекта устройства E3.series.

Автор: number604
Дата: 2026-05-05

Добавьте этот скрипт в Tools -> Customize -> Add-ons (Runtime: pythonw.exe).
При запуске скрипт подключается к текущему экземпляру E3 и последовательно вызывает
`Device.Create2DManufacturingView()` для каждого выделенного устройства в дереве.

Примечание: функция появилась в e3series >= 26.30. Если в вашей установленной
обёртке метод отсутствует, скрипт попытется вызвать метод напрямую на COM-объекте
и выведет подсказку о требуемой версии.
"""

import sys
import traceback

import e3series
import e3series.tools as e3Tools


def main():
    try:
        e3 = e3series.Application()
        job = e3.CreateJobObject()
        if job.GetId() == 0:
            sys.exit("Нет открытого проекта E3.series")

        sys.stdout = e3Tools.E3seriesOutput(e3, sys.stdout)
        print("Create2DManufacturingView — запуск для выделенных устройств в дереве проекта")

        # Попытаться получить выделенные устройства в дереве (несколько fallback)
        cnt, sel_ids = (0, ())
        try:
            cnt, sel_ids = job.GetTreeSelectedAllDeviceIds()
        except Exception:
            try:
                cnt, sel_ids = job.GetTreeSelectedDeviceIds()
            except Exception:
                try:
                    cnt, sel_ids = job.GetSelectedDeviceIds()
                except Exception:
                    cnt, sel_ids = 0, ()

        if not cnt or not sel_ids:
            e3.PutInfo(2, "Не выбрано ни одного устройства в дереве проекта. Выделите устройство и запустите скрипт.")
            return 1

        device = job.CreateDeviceObject()

        for dev_id in sel_ids:
            device.SetId(dev_id)
            dev_name = device.GetName() or f"ID={dev_id}"
            print(f"Обработка устройства: {dev_name} ({dev_id})")

            # Основная логика вызова Create2DManufacturingView
            try:
                if hasattr(device, "Create2DManufacturingView"):
                    print("Вызов wrapper: device.Create2DManufacturingView()")
                    try:
                        res = device.Create2DManufacturingView()
                        print("Результат:", res)
                    except TypeError:
                        # Метод может требовать аргументы — вывести docstring
                        print("Метод требует аргументы. Docstring:")
                        print(device.Create2DManufacturingView.__doc__)
                    except Exception as ex:
                        print("Ошибка при вызове wrapper-метода:", ex)
                        traceback.print_exc()
                else:
                    print("Метод Create2DManufacturingView не найден в обёртке.")
                    # Попробуем вызвать напрямую на COM-объекте
                    if hasattr(device, "_obj") and hasattr(device._obj, "Create2DManufacturingView"):
                        try:
                            res = device._obj.Create2DManufacturingView()
                            print("Результат (через COM):", res)
                        except Exception as ex:
                            print("Ошибка COM-вызова:", ex)
                            traceback.print_exc()
                    else:
                        print("Метод не найден. Требуется e3series>=26.30 или обновление COM/обёртки.")
                        # Для диагностики покажем короткий список методов
                        public_methods = [m for m in dir(device) if not m.startswith('_')]
                        print("Доступные методы (часть):", public_methods[:60])
            except Exception as ex:
                print("Неожиданная ошибка при обработке устройства:", ex)
                traceback.print_exc()

        return 0

    except Exception as exc:
        print("Фатальная ошибка:", exc)
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    sys.exit(main())
