## RadioTowersElevation

## Scope of application 

**RadioTowersElevation** - desktop application, designed for 

**RadioTowersElevation** - десктопное приложение, предназначенное для построения профилей трасс радио-релейных линий связи (РРЛ) и определения оптимальной высоты радиомачт РРЛ.

Program for develop microwave transmission line direct visibility



## Table of contents

  1. [Description](#Description)
  2. [Used technologies](#Used-technologies)
  3. [Installation](#Installation)
  4. [License](#License)

## Description

План:
1. Обьяснение что такое профиль РРЛ и зачем нужен
2. Откуда берётся высота над уровнем моря
3. обьяснение мат модели
4. Демонстрация графического интрефейса
5. возможности предоставляемые пользователю

Радиорелейная линия связи (РРЛ) – это направленный радиоволновой канал связи прямой видимости. Данный канал обладает высокой пропускной способностью и дальностью связи. Применяется для организации различных каналов связи, в том числе и для нужд энергетики. 

Продольный профиль трассы РРЛ описывает естественное искривление поверхности земли, рельеф местности вдоль трассы, высоту и продолжительность объектов на поверхности земли, от которых необходимо отстроить высоту радиомачт, чтоб на пути следования сигнала и в некоторой области вокруг него не было препятствий, препядстувющих передаче сигнала.

![Alt Text](.github/images/way1.jpg)

В построении профилей самой сложной задачей является получение рельефа местности вдоль трассы РРЛ. Для этого использовались [SRTM](https://www2.jpl.nasa.gov/srtm/) данные, которые представляют собой цифровую модель поверхнсти земли описывают высоту над уровнем моря котороя привязана к географическим координатам. Так как в качестве исходных данных для построения профиля используются географические координаты мест установки радиомачт и их высота, был необходим механизм получения географических координат вдоль профиля РРЛ и с конкретным шагом чтобы извлечь значение высоты над уровнем для данной точки из [SRTM](https://www2.jpl.nasa.gov/srtm/). Для этого использовались следующие математические выражения.

![Alt Text](.github/images/azimuth.JPG)

Первоначально определяется растояние между местами установки радиомачт в километрах. Затем определяем азимут, который образуется географическими координатами мест установки радиомачт. Затем составляем массив промеждуточных точек по трассе РРЛ, на основании которого определяем соотвествующие им гегографические координаты по которым составляем массив высот над уровмнем моря.

Далее построение профиля осуществляется по следующим математическим выражениям.

![Alt Text](.github/images/mainmodel.JPG)

Данные выражения сделаны таким образом чтобы сместить центр декардовой системы координат на цент хорды, котороя образуется местами установки радиомачт чтобы упростить расчёт и в последующем построить читаемый профиль. Влияние искривления поверхности земли определяется по радиусу земли (6371.0 км). Высота объектов на поверхности земли задаётся путём построения дополнительной кривой смещённой вверх на заданное растояние. Высота каждой радиомачты задаётся пользователем таким образом чтобы кривая определяющая зону френеля не пересеалось с кривой определяющей высоту обьектов на поверхости земли.

Графический интерфейс реализован при помощи библиотеки [PyQt5](https://pypi.org/project/PyQt5/5.9/) и имеет следующий вид: 

![Alt Text](.github/images/gui.png)

Окно приложения представляют три основные области: таблица для задания параметров радиомачт, таблица для формирования направления трассы РРЛ, профиль РРЛ который строится при при помощи библиотеки [Matplotlib](https://pypi.org/project/matplotlib/2.2.2/). В данном приложении поддреживается построение профилий трасс РРЛ для более 2 радиомачт, что позволяет предоставлять информацию в более наглядном виде. Приложение в качестве сограняемых и открываенмых файлов с исходными данными работает с файлами формата .xlsx, и для их поддержки используется библиотека [Openpyxl](https://pypi.org/project/openpyxl/2.4.8/). Вывод построенных профилей трасс РРЛ реализован путём создания растровых или векторных изображений данных профилей, которые можно гибко настраивать, а именно изменять шрифт и размер текста, шаг сетки графиков, отбражение сложного профиля одним графиком или несколькими, размер и качество изображения.

![Alt Text](.github/images/result.jpg)

## Used technologies

- [Python 3.6.2](https://www.python.org/downloads/) - Python programming language interpreter.
- [Numpy 1.15.0](https://pypi.org/project/numpy/1.15.0/) - general-purpose array-processing package designed to efficiently manipulate large multi-dimensional arrays of arbitrary records without sacrificing too much speed for small multi-dimensional arrays.
- [PyQt5 5.9](https://pypi.org/project/PyQt5/5.9/) - Python binding of the cross-platform GUI toolkit Qt, implemented as a Python plug-in.
- [Matplotlib 2.2.2](https://pypi.org/project/matplotlib/2.2.2/) - library for interactive graphing, scientific publishing, user interface development and web application servers targeting multiple user interfaces and hardcopy output formats.
- [Openpyxl 2.4.8](https://pypi.org/project/openpyxl/2.4.8/) - Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.



## Installation 
Для того чтобы использовать данное приложение необходимо установить компоненты с раздела [Used technologies](#Used-technologies). Первоначально установите интерпретатор Python и Tesseract OCR, а затем при помощи пакетного менеджера *Pip* установите перечисленные пакеты. При применении версий пакетов отличных от предложенных работоспособность приложения не гарантируется.

For using the application necessity to install components from section [Used technologies](#Used-technologies). First of all install Python interpreter and Tesseract OCR, and after that using package manager *Pip* to install listed packages. In case using versions of packages that differ from the proposed, correct work of the application is not ensured.

For use this program you need download SRTM data for your region from https://www2.jpl.nasa.gov/srtm/ or other sources and save it in dirrectory "hgt".


## License 
Licensed under the [MIT](LICENSE.txt) license.	
